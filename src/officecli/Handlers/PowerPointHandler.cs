// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler : IDocumentHandler, Rendering.IRenderModelHost
{
    private readonly PresentationDocument _doc;

    object? Rendering.IRenderModelHost.RenderModel => _doc;
    private readonly string _filePath;
    private HashSet<uint> _usedShapeIds = new();
    private uint _nextShapeId = 10000;
    public int LastFindMatchCount { get; internal set; }
    // Number of elements a no-slash selector Set matched and mutated (Sheet1!row[...]).
    // Read by the CLI/resident to echo the multi-element change count.
    public int LastSelectorSetCount { get; internal set; }

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
    /// trail. Pure Get/Query sessions leave this false.
    /// </summary>
    internal bool Modified { get; set; }

    /// <summary>
    /// Enumerate every <see cref="OpenXmlPart"/> in the package (transitive
    /// walk via the SDK's own <c>GetAllParts</c> extension) yielding each
    /// part's zip-URI (<c>OpenXmlPart.Uri.OriginalString</c>). Used by the
    /// batch emitter's auxiliary-parts scan to surface warnings for parts the
    /// dump surface does not round-trip (tableStyles, viewProps,
    /// handoutMasters, printerSettings, customXml, embedded fonts, tags,
    /// user docProps). Mirrors <see cref="WordHandler.EnumeratePartUris"/>.
    /// </summary>
    internal IEnumerable<string> EnumeratePartUris()
    {
        // Two complementary sources (same rationale as WordHandler):
        //   1. SDK part graph (GetAllParts) — only reachable via the
        //      relationship graph; misses orphan parts but matches every
        //      real pptx file produced by PowerPoint / python-pptx / WPS.
        //   2. Raw zip entries — catches orphan parts dropped by failed
        //      saves and anything the SDK refuses to surface.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in _doc.GetAllParts())
        {
            var u = part.Uri.OriginalString;
            if (seen.Add(u)) yield return u;
        }
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
    /// props vanish). Mirrors <see cref="WordHandler.EnumerateCustomDocPropertyNames"/>.
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

    // Backing FileStream when we open via stream (shared-read mode). null
    // when the package owns its own file handle via PresentationDocument.Open(path).
    private FileStream? _backingStream;
    // Editable sessions open _doc over this in-memory copy of the package so
    // saves swap the file atomically (temp + File.Replace) instead of rewriting
    // it in place. Null for read-only sessions (they never write). See
    // AtomicWriteBack / OfficeCli.Core.AtomicPackageWriter.
    private MemoryStream? _packageStream;

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        // Open via a shared FileStream so external readers (e.g. test harness
        // ZipFile.OpenRead while the handler is alive) don't hit the macOS
        // flock exclusive lock that PresentationDocument.Open(path, editable)
        // would acquire. Editable sessions work against an in-memory copy so the
        // on-disk file is only ever swapped atomically (temp + File.Replace via
        // AtomicWriteBack), never rewritten in place — a process death mid-write
        // otherwise truncates the original with no fallback.
        var share = editable ? FileShare.Read : FileShare.ReadWrite;
        var access = editable ? FileAccess.ReadWrite : FileAccess.Read;
        _backingStream = new FileStream(filePath, FileMode.Open, access, share);
        try
        {
            if (editable)
            {
                _backingStream.Position = 0;
                _packageStream = new MemoryStream();
                _backingStream.CopyTo(_packageStream);
                _packageStream.Position = 0;
            }
            _doc = PresentationDocument.Open((Stream?)_packageStream ?? _backingStream, editable);
            if (editable)
                // fromRawStreams: don't materialize every slide's DOM just to
                // index shape ids — that would make a later Save re-serialize
                // untouched slides (#267). Read ids from the raw part streams.
                InitShapeIdCounter(fromRawStreams: true);
        }
        catch
        {
            _doc?.Dispose();
            _packageStream?.Dispose();
            _backingStream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Get the slide dimensions from the presentation. Falls back to 16:9 (33.867cm × 19.05cm).
    /// </summary>
    private (long width, long height) GetSlideSize()
    {
        var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
        return (sldSz?.Cx?.Value ?? SlideSizeDefaults.Widescreen16x9Cx, sldSz?.Cy?.Value ?? SlideSizeDefaults.Widescreen16x9Cy);
    }

    /// <summary>
    /// Slide size in CSS pixels at 96 DPI — the size a single slide actually
    /// renders at in the HTML preview (EMU ÷ 9525, the canonical EmuConverter
    /// ratio). The screenshot path sizes a single-slide viewport to this so the
    /// PNG is the slide pixel-for-pixel, with no letterbox padding and no scaling.
    /// Physical-derived: a 13.33" and a 10" 16:9 differ in resolution because they
    /// are genuinely different-sized slides. Falls back to 1280×720 (16:9).
    /// </summary>
    internal (int width, int height) GetSlideNativePixels()
    {
        var (w, h) = GetSlideSize();
        if (w <= 0 || h <= 0) return (1280, 720);
        return ((int)Math.Round(w / (double)EmuConverter.EmuPerPx),
                (int)Math.Round(h / (double)EmuConverter.EmuPerPx));
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
            var masters = PowerPointHandler.MastersInOrder(presentationPart);
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"slideMaster[{idx}] not found (total: {masters.Count})");
            return masters[idx - 1].SlideMaster?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }

        var layoutMatch = Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$");
        if (layoutMatch.Success)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = PowerPointHandler.LayoutsInOrder(presentationPart);
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
        Modified = true;
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        if (xpath == null) throw new ArgumentNullException(nameof(xpath));
        if (action == null) throw new ArgumentNullException(nameof(action));
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
            var masters = PowerPointHandler.MastersInOrder(presentationPart);
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
                masters = PowerPointHandler.MastersInOrder(presentationPart);
                if (idx > masters.Count)
                    throw new ArgumentException($"SlideMaster {idx} not found (total: {masters.Count})");
            }
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = PowerPointHandler.LayoutsInOrder(presentationPart);
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
                layouts = PowerPointHandler.LayoutsInOrder(presentationPart);
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
            bool nmCreated = nmPart == null;
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
            // Register <p:notesMasterIdLst> in presentation.xml. AddNewPart only
            // wires the relationship; without the IdLst element the part is an
            // orphan reference and PowerPoint refuses the file with a schema
            // error ("unexpected child element in /p:presentation"). Schema order
            // (CT_Presentation): sldMasterIdLst -> notesMasterIdLst ->
            // handoutMasterIdLst -> sldIdLst -> sldSz -> notesSz -> ...
            if (nmCreated)
            {
                var pres = presentationPart.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                if (pres.NotesMasterIdList == null)
                {
                    var nmRelId = presentationPart.GetIdOfPart(nmPart);
                    var nmIdLst = new NotesMasterIdList(
                        new NotesMasterId { Id = nmRelId });
                    // Insert after sldMasterIdLst if present, else prepend.
                    var sldMasterIdLst = pres.SlideMasterIdList;
                    if (sldMasterIdLst != null)
                        pres.InsertAfter(nmIdLst, sldMasterIdLst);
                    else
                        pres.PrependChild(nmIdLst);
                }
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
        var masters = PowerPointHandler.MastersInOrder(presentationPart);
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
        // targetGlobalIdx is beyond every master's declared range. Real decks
        // can carry MORE layout parts than the master's sldLayoutIdLst declares
        // (rel-linked but undeclared "orphan" layouts — e.g. sample02 declares
        // 1 layout yet ships 11). The emitter enumerates SlideLayoutParts (rels,
        // 11) so raw-set targets /slideLayout[2..11]; grow the shortfall as
        // stub parts under the last master. Unlike the source's orphan form,
        // DECLARE each grown part in sldLayoutIdLst: a replayed slide may bind
        // any of these layouts (the source slide could only ever bind a
        // declared one), and PowerPoint refuses a deck whose slide references
        // a layout missing from its master's sldLayoutIdLst (0x80070570).
        var lastMaster = masters.LastOrDefault();
        if (lastMaster == null) return;
        var idList = lastMaster.SlideMaster?.SlideLayoutIdList;
        if (idList == null && lastMaster.SlideMaster != null)
        {
            idList = new SlideLayoutIdList();
            lastMaster.SlideMaster.SlideLayoutIdList = idList;
        }
        uint nextLayoutId = 2147483649; // min valid sldLayoutId id
        var usedIds = presentationPart.SlideMasterParts
            .Select(m => m.SlideMaster?.SlideLayoutIdList)
            .Where(l => l != null)
            .SelectMany(l => l!.Elements<SlideLayoutId>())
            .Select(e => e.Id?.Value ?? 0)
            .ToHashSet();
        int totalExisting = masters.Sum(m => m.SlideLayoutParts.Count());
        for (int i = totalExisting; i < targetGlobalIdx; i++)
        {
            var extraPart = lastMaster.AddNewPart<SlideLayoutPart>();
            extraPart.SlideLayout = new SlideLayout(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties()))
            ) { Type = SlideLayoutValues.Blank };
            extraPart.SlideLayout.Save();
            extraPart.AddPart(lastMaster);
            if (idList != null)
            {
                while (usedIds.Contains(nextLayoutId)) nextLayoutId++;
                idList.AppendChild(new SlideLayoutId
                {
                    Id = nextLayoutId,
                    RelationshipId = lastMaster.GetIdOfPart(extraPart),
                });
                usedIds.Add(nextLayoutId);
            }
        }
        if (lastMaster.SlideMaster != null) lastMaster.SlideMaster.Save();
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

            // Every SlideMasterPart must reference a ThemePart. Give each grown
            // master its OWN distinct (placeholder) theme part rather than sharing
            // the presentation's primary theme: `add-part theme` overwrites it
            // with the source's per-master theme content, and a later
            // delete-of-its-own-theme must not orphan a theme another master
            // shares. Seed minimal valid theme content (replaced on replay when
            // the deck carries per-master themes).
            var grownTheme = newPart.AddNewPart<ThemePart>();
            if (presentationPart.ThemePart?.Theme != null)
                grownTheme.Theme = (DocumentFormat.OpenXml.Drawing.Theme)presentationPart.ThemePart.Theme.CloneNode(true);
            else
                grownTheme.Theme = new DocumentFormat.OpenXml.Drawing.Theme();
            grownTheme.Theme.Save();

            sldMasterIdLst.AppendChild(new SlideMasterId { Id = nextId++, RelationshipId = rId });
        }
        presentation.Save();
    }

    // PowerPoint requires every <p:sldMasterId>/@id AND every
    // <p:sldLayoutId>/@id (across all masters) to be unique within one shared
    // id space — PowerPoint's own decks number them as a single monotonic run
    // (master=N, its layouts=N+1.., next master continues after the last
    // layout). The SDK and our GrowSlideMasterParts pick master ids starting at
    // 2147483648 while only avoiding existing *master* ids, so a grown master
    // can land on a value a source master's sldLayoutIdLst already uses
    // (e.g. master2 id 2147483649 == master1 layout1 id 2147483649). The SDK
    // schema-validates this duplicate fine, but PowerPoint rejects the package
    // with 0x80070570 ("file or directory corrupted"). Reconcile at save: keep
    // layout ids fixed (slides never reference them), reassign only colliding
    // master ids to fresh unused values. Runs once per save cycle, after every
    // raw-set has landed the masters' real sldLayoutIdLst.
    private void ReconcileSlideMasterIds()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var sldMasterIdLst = presentation?.SlideMasterIdList;
        if (presentation == null || sldMasterIdLst == null) return;

        // Collect every layout id (these stay put) plus seed with the values
        // we will keep for masters as we walk them.
        var used = new HashSet<uint>();
        foreach (var mp in _doc.PresentationPart!.SlideMasterParts)
        {
            var layoutIds = mp.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>();
            if (layoutIds == null) continue;
            foreach (var lid in layoutIds)
                if (lid.Id?.Value is uint v) used.Add(v);
        }

        bool changed = false;
        foreach (var smId in sldMasterIdLst.Elements<SlideMasterId>())
        {
            uint id = smId.Id?.Value ?? 2147483648u;
            if (used.Contains(id))
            {
                uint repl = 2147483648u;
                while (used.Contains(repl)) repl++;
                smId.Id = repl;
                id = repl;
                changed = true;
            }
            used.Add(id);
        }

        if (changed) presentation.Save();
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

            case "smartart":
                // SmartArt graphicFrame references four separate OOXML
                // sub-parts under the owning SlidePart: data (dgm:dataModel),
                // layout (dgm:layoutDef), colors (dgm:colorsDef), style
                // (dgm:styleDef). The graphicFrame's <dgm:relIds> carries the
                // four rIds, so dump→batch→replay byte-equality requires that
                // the rIds match the source's. Callers MAY pass explicit rIds
                // via properties {data, layout, colors, quickStyle}; when
                // omitted the SDK allocates fresh ones. Each part is seeded
                // with a minimal typed root so subsequent raw-set replace
                // ops can target /dgm:dataModel etc.
                var saSlideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!saSlideMatch.Success)
                    throw new ArgumentException(
                        "SmartArt must be added under a slide: add-part <file> '/slide[N]' --type smartart");
                var saSlideIdx = int.Parse(saSlideMatch.Groups[1].Value);
                var saSlideParts = GetSlideParts().ToList();
                if (saSlideIdx < 1 || saSlideIdx > saSlideParts.Count)
                    throw new ArgumentException($"Slide index {saSlideIdx} out of range");
                var saSlidePart = saSlideParts[saSlideIdx - 1];

                string? dataRid     = properties != null && properties.TryGetValue("data", out var dv) ? dv : null;
                string? layoutRid   = properties != null && properties.TryGetValue("layout", out var lv) ? lv : null;
                string? colorsRid   = properties != null && properties.TryGetValue("colors", out var cv) ? cv : null;
                string? qsRid       = properties != null && properties.TryGetValue("quickStyle", out var qv) ? qv : null;

                // Inline diagram part content. The dump emitter carries each
                // sub-part's verbatim XML here so add-part writes it directly
                // into the freshly-created part. This supersedes the legacy
                // "create seed + separate raw-set replace" flow: the SDK
                // allocates the diagram parts under /ppt/graphics/dataN.xml
                // (its own naming base, N incrementing package-globally),
                // NOT the source's /ppt/diagrams/data1.xml, so a raw-set
                // pre-targeted at the source URI never resolved (FindPartByZipUri
                // miss) and the parts persisted EMPTY → blank/broken SmartArt.
                // Writing content at creation time is URI-agnostic and robust.
                string? dataXml     = properties != null && properties.TryGetValue("dataXml", out var dxv) ? dxv : null;
                string? layoutXml   = properties != null && properties.TryGetValue("layoutXml", out var lxv) ? lxv : null;
                string? colorsXml   = properties != null && properties.TryGetValue("colorsXml", out var cxv) ? cxv : null;
                string? qsXml       = properties != null && properties.TryGetValue("quickStyleXml", out var qxv) ? qxv : null;
                string? drawingXml  = properties != null && properties.TryGetValue("drawingXml", out var drxv) ? drxv : null;
                string? drawingRelId = properties != null && properties.TryGetValue("drawingRelId", out var drrv) ? drrv : null;

                // Reuse-aware creation: a second SmartArt on the same slide may
                // share the layout/colors/quickStyle parts (see GetOrAddPinnedPart).
                // Data parts are per-diagram unique, but the helper handles all
                // four uniformly and skips rewriting a reused part's content.
                DiagramDataPart   dataPart   = GetOrAddPinnedPart<DiagramDataPart>(saSlidePart, dataRid, out var dataCreated);
                DiagramLayoutDefinitionPart layoutPart = GetOrAddPinnedPart<DiagramLayoutDefinitionPart>(saSlidePart, layoutRid, out var layoutCreated);
                DiagramColorsPart colorsPart = GetOrAddPinnedPart<DiagramColorsPart>(saSlidePart, colorsRid, out var colorsCreated);
                DiagramStylePart  stylePart  = GetOrAddPinnedPart<DiagramStylePart>(saSlidePart, qsRid, out var styleCreated);

                // Write the real content when supplied; else seed a minimal
                // typed root (keeps direct CLI `add-part smartart` usable).
                // Skip parts reused from a prior SmartArt — their content is
                // already written and the second op may not carry it.
                if (dataCreated)
                {
                    WriteDiagramPartXml(dataPart, dataXml, () =>
                        new DocumentFormat.OpenXml.Drawing.Diagrams.DataModelRoot(
                            new DocumentFormat.OpenXml.Drawing.Diagrams.PointList(),
                            new DocumentFormat.OpenXml.Drawing.Diagrams.ConnectionList()));
                    // Pictures embedded in the diagram are referenced from the data
                    // part's own .rels; re-attach them with pinned rIds.
                    AttachDiagramImages(dataPart, properties, "dataImage");
                    // External hyperlinks on diagram nodes (<a:hlinkClick r:id>) live
                    // on the data part's own .rels; re-add with pinned rIds.
                    AttachDiagramHyperlinks(dataPart, properties, "dataHlink");
                }
                if (layoutCreated)
                    WriteDiagramPartXml(layoutPart, layoutXml,
                        () => new DocumentFormat.OpenXml.Drawing.Diagrams.LayoutDefinition());
                if (colorsCreated)
                    WriteDiagramPartXml(colorsPart, colorsXml,
                        () => new DocumentFormat.OpenXml.Drawing.Diagrams.ColorsDefinition());
                if (styleCreated)
                    WriteDiagramPartXml(stylePart, qsXml,
                        () => new DocumentFormat.OpenXml.Drawing.Diagrams.StyleDefinition());

                // The DSP cached-drawing part is referenced from the data XML
                // via <dsp:dataModelExt relId="...">. That relId resolves
                // against the SLIDE part's relationships (the drawing part is
                // a slide-level part of type .../2007/relationships/diagramDrawing,
                // sibling to the data/layout/colors/qs rels — NOT a child of
                // the data part). Create it on saSlidePart with the pinned
                // relId so the reference resolves; otherwise PowerPoint
                // refuses the file (0x80070570).
                if (!string.IsNullOrEmpty(drawingXml) && !string.IsNullOrEmpty(drawingRelId))
                {
                    var existingDrawing = saSlidePart.Parts
                        .FirstOrDefault(p => p.RelationshipId == drawingRelId).OpenXmlPart;
                    if (existingDrawing is not DiagramPersistLayoutPart)
                    {
                        var drawingPart = saSlidePart.AddNewPart<DiagramPersistLayoutPart>(
                            "application/vnd.ms-office.drawingml.diagramDrawing+xml", drawingRelId);
                        // drawingXml is always present on this branch; the seed
                        // fallback is unreachable here but supplied for the typed
                        // signature.
                        WriteDiagramPartXml(drawingPart, drawingXml,
                            () => new DocumentFormat.OpenXml.Drawing.Diagrams.DataModelRoot());
                        // The DSP cached drawing re-references the same pictures for
                        // rendering via its own .rels; re-attach with pinned rIds.
                        AttachDiagramImages(drawingPart, properties, "drawingImage");
                        AttachDiagramHyperlinks(drawingPart, properties, "drawingHlink");
                    }
                }

                // Encode all four rIds in the RelId field — callers (batch
                // emit / replay) need to know each part's id to write the
                // matching dgm:relIds on the graphicFrame. Format:
                // "data=rIdX;layout=rIdY;colors=rIdZ;quickStyle=rIdW".
                var dataActualRid   = saSlidePart.GetIdOfPart(dataPart);
                var layoutActualRid = saSlidePart.GetIdOfPart(layoutPart);
                var colorsActualRid = saSlidePart.GetIdOfPart(colorsPart);
                var styleActualRid  = saSlidePart.GetIdOfPart(stylePart);

                // Inject a minimal <p:graphicFrame> into the slide's spTree
                // so GetSmartArtsOnSlide (which finds SmartArt only via the
                // graphicFrame + dgm:relIds anchor) can see this SmartArt on
                // the next dump. Without the host frame, a SmartArt created
                // by direct `add-part smartart` is silently dropped on dump.
                // The dump-emitter path passes `skip-frame=true` because it
                // raw-set appends the source's full graphicFrame (with real
                // position/size/name) immediately after — otherwise we'd
                // emit a stub-plus-real pair on every replay.
                var skipFrame = properties != null
                    && properties.TryGetValue("skip-frame", out var sf)
                    && (sf == "true" || sf == "1");
                if (!skipFrame)
                {
                    var saSlide = GetSlide(saSlidePart);
                    var saSpTree = saSlide.CommonSlideData?.ShapeTree;
                    if (saSpTree != null)
                    {
                        // Allocate a non-colliding cNvPr id within the slide.
                        // R42-T1: spTree carries pptx-namespaced <p:nvSpPr>/<p:nvGrpSpPr>/<p:nvGraphicFramePr>
                        // wrappers whose cNvPr maps to DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties,
                        // NOT the Drawing-namespace type. The wrong SDK type matched zero descendants,
                        // pinning nextId at 1 and colliding with nvGrpSpPr cNvPr id=1.
                        uint nextId = 1;
                        foreach (var nv in saSpTree.Descendants<DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties>())
                        {
                            if (nv.Id?.Value >= nextId) nextId = nv.Id!.Value + 1;
                        }
                        // CT_GraphicalObjectFrame: nvGraphicFramePr / xfrm /
                        // graphic. xfrm carries a default 6"x4.5" host area
                        // anchored at (1in, 1in) — PowerPoint will rescale
                        // on first render anyway; the values just keep the
                        // shape selectable.
                        const long EmuIn = 914400;
                        var gf = new DocumentFormat.OpenXml.Presentation.GraphicFrame(
                            new DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameProperties(
                                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = nextId, Name = $"Diagram {nextId}" },
                                new DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameDrawingProperties(
                                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
                                new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                            new DocumentFormat.OpenXml.Presentation.Transform(
                                new DocumentFormat.OpenXml.Drawing.Offset { X = 1 * EmuIn, Y = 1 * EmuIn },
                                new DocumentFormat.OpenXml.Drawing.Extents { Cx = 6 * EmuIn, Cy = (long)(4.5 * EmuIn) }),
                            new DocumentFormat.OpenXml.Drawing.Graphic(
                                new DocumentFormat.OpenXml.Drawing.GraphicData(
                                    new DocumentFormat.OpenXml.Drawing.Diagrams.RelationshipIds
                                    {
                                        DataPart = dataActualRid,
                                        LayoutPart = layoutActualRid,
                                        ColorPart = colorsActualRid,
                                        StylePart = styleActualRid,
                                    })
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" }));
                        saSpTree.AppendChild(gf);
                        saSlide.Save();
                    }
                }

                var encoded = $"data={dataActualRid};layout={layoutActualRid};colors={colorsActualRid};quickStyle={styleActualRid}";
                return (encoded, parentPartPath);

            case "video":
            case "audio":
                // Phase 3c-media. Mirror Phase 3b SmartArt: create the
                // underlying parts (MediaDataPart + ImagePart thumbnail)
                // and pin all three rIds via properties so the post-replay
                // <p:pic> appended by raw-set finds the same rIds it carried
                // in the source. The graphicFrame analogue here is the
                // <p:pic> referencing <a:videoFile r:link=…/> +
                // <p14:media r:embed=…/> in nvPr, plus <a:blip r:embed=…/>
                // in blipFill for the thumbnail.
                //
                // Props (all optional except data + thumbnail-data):
                //   data                   = base64 binary (mp4/m4a/…)
                //   content-type           = "video/mp4" / "audio/mpeg" / …
                //   extension              = ".mp4" / ".m4a" (for the
                //                            MediaDataPart URI extension;
                //                            best-effort from content-type
                //                            when omitted)
                //   thumbnail-data         = base64 image binary
                //   thumbnail-content-type = "image/png" / "image/jpeg"
                //   video-rid / audio-rid  = pinned VideoReference / AudioReference rId
                //   media-rid              = pinned p14:media MediaReference rId
                //   thumbnail-rid          = pinned ImagePart rId
                //
                // Audio uses AddAudioReferenceRelationship; video uses
                // AddVideoReferenceRelationship. Both ALSO add a
                // MediaReferenceRelationship (the p14:media r:embed is
                // distinct from the legacy r:link).
                var mediaSlideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!mediaSlideMatch.Success)
                    throw new ArgumentException(
                        $"{partType} must be added under a slide: add-part <file> '/slide[N]' --type {partType}");
                var mediaSlideIdx = int.Parse(mediaSlideMatch.Groups[1].Value);
                var mediaSlidePartsList = GetSlideParts().ToList();
                if (mediaSlideIdx < 1 || mediaSlideIdx > mediaSlidePartsList.Count)
                    throw new ArgumentException($"Slide index {mediaSlideIdx} out of range");
                var mediaSlidePart = mediaSlidePartsList[mediaSlideIdx - 1];

                if (properties == null || !properties.TryGetValue("data", out var mediaB64) || string.IsNullOrEmpty(mediaB64))
                    throw new ArgumentException(
                        $"add-part {partType} requires property 'data' (base64 binary)");
                byte[] mediaBytes;
                try { mediaBytes = Convert.FromBase64String(mediaB64); }
                catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'data' is not valid base64"); }

                var mediaContentType = properties.TryGetValue("content-type", out var mct) && !string.IsNullOrEmpty(mct)
                    ? mct
                    : (partType == "video" ? "video/mp4" : "audio/mpeg");
                var mediaExt = properties.TryGetValue("extension", out var mxt) && !string.IsNullOrEmpty(mxt)
                    ? mxt
                    : mediaContentType switch {
                        "video/mp4" => ".mp4", "video/x-msvideo" => ".avi",
                        "video/x-ms-wmv" => ".wmv", "video/mpeg" => ".mpg",
                        "video/quicktime" => ".mov",
                        "audio/mpeg" => ".mp3", "audio/wav" => ".wav",
                        "audio/x-ms-wma" => ".wma", "audio/mp4" => ".m4a",
                        _ => ".bin" };

                var mediaDataPart = _doc.CreateMediaDataPart(mediaContentType, mediaExt);
                using (var inStream = new MemoryStream(mediaBytes))
                    mediaDataPart.FeedData(inStream);

                string? pinnedVideoRid = properties.TryGetValue("video-rid", out var vr) ? vr : null;
                string? pinnedAudioRid = properties.TryGetValue("audio-rid", out var ar) ? ar : null;

                // rel-only: a timing-tree / transition SOUND (<p:sndTgt r:embed>,
                // <p:snd r:embed>) needs ONLY a bare `.../audio` relationship to a
                // MediaDataPart with the source rId — NOT the <p:pic> media
                // machinery (MediaReference + thumbnail). Without this the audio
                // rId in the raw-passed-through timing tree dangled and PowerPoint
                // refused the deck (0x80070570). Create the audio rel with the
                // pinned rId and return; the timing slice already references it.
                if (partType == "audio"
                    && properties.TryGetValue("rel-only", out var relOnly) && IsTruthy(relOnly))
                {
                    var soundRid = !string.IsNullOrEmpty(pinnedAudioRid)
                        ? mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart, pinnedAudioRid).Id
                        : mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart).Id;
                    return ($"audio={soundRid}", parentPartPath);
                }
                string? pinnedMediaRid = properties.TryGetValue("media-rid", out var mr) ? mr : null;
                string? pinnedThumbRid = properties.TryGetValue("thumbnail-rid", out var tr) ? tr : null;

                string linkRelId;
                if (partType == "video")
                {
                    linkRelId = !string.IsNullOrEmpty(pinnedVideoRid)
                        ? mediaSlidePart.AddVideoReferenceRelationship(mediaDataPart, pinnedVideoRid).Id
                        : mediaSlidePart.AddVideoReferenceRelationship(mediaDataPart).Id;
                }
                else
                {
                    linkRelId = !string.IsNullOrEmpty(pinnedAudioRid)
                        ? mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart, pinnedAudioRid).Id
                        : mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart).Id;
                }
                var mediaEmbedRid = !string.IsNullOrEmpty(pinnedMediaRid)
                    ? mediaSlidePart.AddMediaReferenceRelationship(mediaDataPart, pinnedMediaRid).Id
                    : mediaSlidePart.AddMediaReferenceRelationship(mediaDataPart).Id;

                // Thumbnail (poster) — required so the <a:blip r:embed> in
                // the <p:pic>'s blipFill resolves on replay. Caller MAY
                // omit thumbnail-data; we then seed a 1x1 transparent PNG
                // (mirrors the AddMedia helper's placeholder path).
                byte[] thumbBytes;
                PartTypeInfo thumbType;
                if (properties.TryGetValue("thumbnail-data", out var tdB64) && !string.IsNullOrEmpty(tdB64))
                {
                    try { thumbBytes = Convert.FromBase64String(tdB64); }
                    catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'thumbnail-data' is not valid base64"); }
                    var thumbCT = properties.TryGetValue("thumbnail-content-type", out var tct) && !string.IsNullOrEmpty(tct)
                        ? tct
                        : "image/png";
                    thumbType = thumbCT switch {
                        "image/png" => ImagePartType.Png, "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif, "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" or "image/tif" => ImagePartType.Tiff, _ => ImagePartType.Png };
                }
                else
                {
                    thumbBytes = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    thumbType = ImagePartType.Png;
                }
                var thumbImagePart = !string.IsNullOrEmpty(pinnedThumbRid)
                    ? mediaSlidePart.AddImagePart(thumbType, pinnedThumbRid)
                    : mediaSlidePart.AddImagePart(thumbType);
                using (var thumbStream = new MemoryStream(thumbBytes))
                    thumbImagePart.FeedData(thumbStream);
                var thumbActualRid = mediaSlidePart.GetIdOfPart(thumbImagePart);

                // Encode three rIds — emitter / replay caller may use any
                // of them when writing the <p:pic> XML via raw-set. The
                // (RelId, PartPath) tuple's RelId is consumed by callers
                // who do their own bookkeeping; format mirrors smartart.
                var mediaKey = partType == "video" ? "video" : "audio";
                var encodedMedia = $"{mediaKey}={linkRelId};media={mediaEmbedRid};thumbnail={thumbActualRid}";
                return (encodedMedia, parentPartPath);

            case "model3d":
            case "3dmodel":
                // Phase 3c-3d. Mirrors Phase 3c-media (video/audio). The
                // PPT 3D model lives inside <mc:AlternateContent>:
                //   <mc:Choice Requires="am3d">
                //     <p:graphicFrame>... <am3d:model3d r:embed=…>
                //         <am3d:raster><am3d:blip r:embed=…/></am3d:raster>
                //   <mc:Fallback><p:pic>... <a:blip r:embed=…/></p:pic>
                // Two rels back the slice:
                //  - relType .../office/2017/06/relationships/model3d
                //    -> the .glb binary (created via AddExtendedPart;
                //       SDK lacks a typed Model3DPart)
                //  - relType .../officeDocument/2006/relationships/image
                //    -> a static thumbnail PNG (shared by am3d:raster's
                //       blip AND the Fallback p:pic's blipFill)
                //
                // Props (all optional except data + thumbnail-data):
                //   data                   = base64 .glb bytes
                //   content-type           = "model/gltf.binary" (default)
                //   extension              = ".glb" (default)
                //   model3d-rid            = pinned model3d ExtendedPart rId
                //   thumbnail-data         = base64 thumbnail image bytes
                //   thumbnail-content-type = "image/png" (default) / "image/jpeg"
                //   thumbnail-rid          = pinned thumbnail ImagePart rId
                //
                // Returned encoded relId: "model3d=rIdA;thumbnail=rIdB".
                // No shape XML is inserted under the slide; the companion
                // raw-set append carries the full <mc:AlternateContent>
                // verbatim with the matching pinned rIds.
                var m3dSlideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!m3dSlideMatch.Success)
                    throw new ArgumentException(
                        $"{partType} must be added under a slide: add-part <file> '/slide[N]' --type {partType}");
                var m3dSlideIdx = int.Parse(m3dSlideMatch.Groups[1].Value);
                var m3dSlideParts = GetSlideParts().ToList();
                if (m3dSlideIdx < 1 || m3dSlideIdx > m3dSlideParts.Count)
                    throw new ArgumentException($"Slide index {m3dSlideIdx} out of range");
                var m3dSlidePart = m3dSlideParts[m3dSlideIdx - 1];

                // 'data' may be absent ONLY when callers (rare) are pre-
                // declaring rels with empty parts; the dump emitter always
                // passes it.
                if (properties == null || !properties.TryGetValue("data", out var m3dB64) || string.IsNullOrEmpty(m3dB64))
                    throw new ArgumentException(
                        $"add-part {partType} requires property 'data' (base64 .glb bytes)");
                byte[] m3dBytes;
                try { m3dBytes = Convert.FromBase64String(m3dB64); }
                catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'data' is not valid base64"); }

                var m3dContentType = properties.TryGetValue("content-type", out var m3dct) && !string.IsNullOrEmpty(m3dct)
                    ? m3dct
                    : "model/gltf.binary";
                var m3dExt = properties.TryGetValue("extension", out var m3dxt) && !string.IsNullOrEmpty(m3dxt)
                    ? (m3dxt.StartsWith('.') ? m3dxt : "." + m3dxt)
                    : ".glb";

                string? pinnedM3dRid   = properties.TryGetValue("model3d-rid", out var m3dr) ? m3dr : null;
                string? pinnedM3dThumb = properties.TryGetValue("thumbnail-rid", out var m3dtr) ? m3dtr : null;

                const string m3dRelType = "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
                var m3dPart = !string.IsNullOrEmpty(pinnedM3dRid)
                    ? m3dSlidePart.AddExtendedPart(m3dRelType, m3dContentType, m3dExt, pinnedM3dRid)
                    : m3dSlidePart.AddExtendedPart(m3dRelType, m3dContentType, m3dExt);
                using (var s = new MemoryStream(m3dBytes)) m3dPart.FeedData(s);
                var m3dActualRid = m3dSlidePart.GetIdOfPart(m3dPart);

                // Thumbnail (am3d:raster blip + Fallback blipFill share one
                // ImagePart). Caller may omit thumbnail-data; we then seed
                // the same 1x1 transparent PNG used by the video/audio
                // placeholder path.
                byte[] m3dThumbBytes;
                PartTypeInfo m3dThumbType;
                if (properties.TryGetValue("thumbnail-data", out var m3dtdB64) && !string.IsNullOrEmpty(m3dtdB64))
                {
                    try { m3dThumbBytes = Convert.FromBase64String(m3dtdB64); }
                    catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'thumbnail-data' is not valid base64"); }
                    var m3dThumbCT = properties.TryGetValue("thumbnail-content-type", out var m3dtct) && !string.IsNullOrEmpty(m3dtct)
                        ? m3dtct
                        : "image/png";
                    m3dThumbType = m3dThumbCT switch {
                        "image/png" => ImagePartType.Png, "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif, "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" or "image/tif" => ImagePartType.Tiff, _ => ImagePartType.Png };
                }
                else
                {
                    m3dThumbBytes = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    m3dThumbType = ImagePartType.Png;
                }
                var m3dThumbPart = !string.IsNullOrEmpty(pinnedM3dThumb)
                    ? m3dSlidePart.AddImagePart(m3dThumbType, pinnedM3dThumb)
                    : m3dSlidePart.AddImagePart(m3dThumbType);
                using (var s = new MemoryStream(m3dThumbBytes)) m3dThumbPart.FeedData(s);
                var m3dThumbActualRid = m3dSlidePart.GetIdOfPart(m3dThumbPart);

                var encodedM3d = $"model3d={m3dActualRid};thumbnail={m3dThumbActualRid}";
                return (encodedM3d, parentPartPath);

            case "ole":
            case "oleobject":
                // Phase 3c-ole. Mirrors Phase 3c-media (video/audio) and
                // Phase 3c-3d. PPT OLE embed lives inside a <p:graphicFrame>:
                //   <a:graphicData uri=".../presentationml/2006/ole">
                //     <p:oleObj progId="…" showAsIcon="1" r:id="rIdA" …>
                //       <p:embed/>
                //       <p:pic>... <a:blip r:embed="rIdB"/> ...</p:pic>
                //     </p:oleObj>
                // Two rels back the slice:
                //  - relType .../officeDocument/2006/relationships/oleObject
                //    -> the embedded payload. EmbeddedPackagePart for OOXML
                //       containers (.xlsx/.docx/.pptx + macro/template
                //       siblings, content type vnd.openxmlformats-…), else
                //       EmbeddedObjectPart for generic binaries (.bin OLE10,
                //       legacy .doc/.xls, PDF, etc.).
                //  - relType .../officeDocument/2006/relationships/image
                //    -> the icon/thumbnail ImagePart for the inner <p:pic>.
                //
                // Props (all optional except data + thumbnail-data):
                //   data                   = base64 OLE payload bytes
                //   content-type           = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                //                            (for .xlsx), or any package /
                //                            generic OLE content-type. Drives
                //                            the Package vs Object auto-select.
                //   extension              = ".xlsx" / ".docx" / ".bin" (drives
                //                            the part URI extension when we
                //                            fall through to EmbeddedObjectPart;
                //                            EmbeddedPackagePart picks its own
                //                            extension from the PartTypeInfo).
                //   ole-rid                = pinned OLE part rId
                //   thumbnail-data         = base64 icon image bytes
                //   thumbnail-content-type = "image/png" (default) / "image/jpeg"
                //   thumbnail-rid          = pinned thumbnail ImagePart rId
                //
                // Returned encoded relId: "ole=rIdA;thumbnail=rIdB".
                // No shape XML is inserted under the slide; the companion
                // raw-set append carries the full <p:graphicFrame> verbatim
                // with the matching pinned rIds.
                var oleSlideMatchAP = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!oleSlideMatchAP.Success)
                    throw new ArgumentException(
                        $"{partType} must be added under a slide: add-part <file> '/slide[N]' --type {partType}");
                var oleSlideIdxAP = int.Parse(oleSlideMatchAP.Groups[1].Value);
                var oleSlidePartsAP = GetSlideParts().ToList();
                if (oleSlideIdxAP < 1 || oleSlideIdxAP > oleSlidePartsAP.Count)
                    throw new ArgumentException($"Slide index {oleSlideIdxAP} out of range");
                var oleSlidePartAP = oleSlidePartsAP[oleSlideIdxAP - 1];

                if (properties == null || !properties.TryGetValue("data", out var oleB64) || string.IsNullOrEmpty(oleB64))
                    throw new ArgumentException(
                        $"add-part {partType} requires property 'data' (base64 OLE payload bytes)");
                byte[] oleBytes;
                try { oleBytes = Convert.FromBase64String(oleB64); }
                catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'data' is not valid base64"); }

                var oleContentTypeAP = properties.TryGetValue("content-type", out var olct) && !string.IsNullOrEmpty(olct)
                    ? olct
                    : "application/vnd.openxmlformats-officedocument.oleObject";
                var oleExtAP = properties.TryGetValue("extension", out var olxt) && !string.IsNullOrEmpty(olxt)
                    ? (olxt.StartsWith('.') ? olxt : "." + olxt)
                    : ".bin";

                string? pinnedOleRid   = properties.TryGetValue("ole-rid", out var olr) ? olr : null;
                string? pinnedOleThumb = properties.TryGetValue("thumbnail-rid", out var oltr) ? oltr : null;

                // Auto-select EmbeddedPackagePart vs EmbeddedObjectPart by
                // content-type. The OOXML package family carries content
                // types starting with "application/vnd.openxmlformats-".
                // Everything else (oleObject generic, application/pdf,
                // application/octet-stream, vnd.ms-excel for legacy .xls,
                // etc.) goes through EmbeddedObjectPart with its raw
                // content-type preserved on the part. EmbeddedPackagePart
                // requires a typed PartTypeInfo (EmbeddedPackagePartType.*);
                // map extension → typed value via OleHelper, fall back to
                // EmbeddedObjectPart if the extension is unrecognized.
                OpenXmlPart olePart;
                PartTypeInfo? packagePti = null;
                bool oleIsPackage = oleContentTypeAP.StartsWith(
                    "application/vnd.openxmlformats-officedocument.",
                    StringComparison.OrdinalIgnoreCase)
                    && !oleContentTypeAP.Equals(
                        "application/vnd.openxmlformats-officedocument.oleObject",
                        StringComparison.OrdinalIgnoreCase);
                if (oleIsPackage)
                {
                    // GetPackagePartTypeInfo takes a path; we feed a synthetic
                    // path with the extension. Returns null for unknown exts
                    // — in that case route to EmbeddedObjectPart so the bytes
                    // still survive (with a generic part type).
                    packagePti = OfficeCli.Core.OleHelper.GetPackagePartTypeInfo("x" + oleExtAP);
                    if (packagePti == null) oleIsPackage = false;
                }
                if (oleIsPackage)
                {
                    olePart = !string.IsNullOrEmpty(pinnedOleRid)
                        ? oleSlidePartAP.AddEmbeddedPackagePart(packagePti!.Value, pinnedOleRid)
                        : oleSlidePartAP.AddEmbeddedPackagePart(packagePti!.Value);
                }
                else
                {
                    olePart = !string.IsNullOrEmpty(pinnedOleRid)
                        ? oleSlidePartAP.AddEmbeddedObjectPart(oleContentTypeAP, pinnedOleRid)
                        : oleSlidePartAP.AddEmbeddedObjectPart(oleContentTypeAP);
                }
                using (var s = new MemoryStream(oleBytes)) olePart.FeedData(s);
                var oleActualRid = oleSlidePartAP.GetIdOfPart(olePart);

                // Thumbnail icon image. Caller may omit thumbnail-data; we
                // then seed the same 1x1 transparent PNG used by video/audio
                // and the model3d placeholder paths.
                byte[] oleThumbBytes;
                PartTypeInfo oleThumbType;
                if (properties.TryGetValue("thumbnail-data", out var oltdB64) && !string.IsNullOrEmpty(oltdB64))
                {
                    try { oleThumbBytes = Convert.FromBase64String(oltdB64); }
                    catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'thumbnail-data' is not valid base64"); }
                    var oleThumbCT = properties.TryGetValue("thumbnail-content-type", out var oltct) && !string.IsNullOrEmpty(oltct)
                        ? oltct
                        : "image/png";
                    oleThumbType = oleThumbCT switch {
                        "image/png" => ImagePartType.Png, "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif, "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" or "image/tif" => ImagePartType.Tiff,
                        "image/x-emf" => ImagePartType.Emf, "image/x-wmf" => ImagePartType.Wmf,
                        _ => ImagePartType.Png };
                }
                else
                {
                    oleThumbBytes = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    oleThumbType = ImagePartType.Png;
                }
                var oleThumbPart = !string.IsNullOrEmpty(pinnedOleThumb)
                    ? oleSlidePartAP.AddImagePart(oleThumbType, pinnedOleThumb)
                    : oleSlidePartAP.AddImagePart(oleThumbType);
                using (var s = new MemoryStream(oleThumbBytes)) oleThumbPart.FeedData(s);
                var oleThumbActualRid = oleSlidePartAP.GetIdOfPart(oleThumbPart);

                var encodedOle = $"ole={oleActualRid};thumbnail={oleThumbActualRid}";
                return (encodedOle, parentPartPath);

            case "image":
                // Generic ImagePart attached to a slide / slideMaster / slideLayout
                // / notesMaster. Used by dump→replay to round-trip image
                // references that live in raw-set'd master/layout XML — the
                // master raw XML carries `r:embed="rIdN"` on <p:pic> blipFills,
                // but the underlying ImagePart is enumerated separately by the
                // SDK and was never re-emitted, so post-replay validate flagged
                // "rIdN does not exist" on slideMaster2 etc.
                //
                // Required properties:
                //   data         = base64 image bytes
                // Optional properties:
                //   content-type = "image/png" / "image/jpeg" / "image/gif" / …
                //                  (default "image/png")
                //   rid          = pinned relationship id; when omitted the SDK
                //                  allocates one
                OpenXmlPart imageHost = parentPartPath switch
                {
                    "/" => presentationPart,
                    "/notesMaster" => (OpenXmlPart?)presentationPart.NotesMasterPart
                        ?? throw new ArgumentException("add-part image /notesMaster: no notes master in deck"),
                    // A theme's <a:fmtScheme>/<a:fillStyleLst>/<a:blipFill> can
                    // reference a texture image via r:embed; the raw-set'd theme
                    // XML carries that reference but the ImagePart lives in the
                    // theme's own .rels, enumerated separately. Without carrying
                    // it the rId dangles and PowerPoint refuses to open the deck.
                    "/theme" => (OpenXmlPart?)presentationPart.ThemePart
                        ?? throw new ArgumentException("add-part image /theme: presentation has no theme part"),
                    _ => null!,
                };
                if (imageHost == null)
                {
                    var smMatch = System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                    var slMatch = smMatch.Success ? null : System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                    var sldMatch = (smMatch.Success || (slMatch?.Success ?? false)) ? null : System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                    // CONSISTENCY(notes-image-host): mirror /slide[N]/master/layout
                    // entries — dump emits add-part image /noteSlide[N] BEFORE the
                    // notes raw-set replace so r:embed references in the notesSlide
                    // XML resolve to a real ImagePart on replay. Without this the
                    // post-replay notesSlide carries a dangling rId and PowerPoint
                    // renders the speaker-notes picture as a broken placeholder.
                    var nsMatch = (smMatch.Success || (slMatch?.Success ?? false) || (sldMatch?.Success ?? false))
                        ? null
                        : System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/noteSlide\[(\d+)\]$");
                    if (smMatch.Success)
                    {
                        var smIdx = int.Parse(smMatch.Groups[1].Value);
                        var smParts = PowerPointHandler.MastersInOrder(presentationPart);
                        if (smIdx < 1) throw new ArgumentException($"slideMaster index {smIdx} out of range");
                        if (smIdx > smParts.Count)
                        {
                            // CONSISTENCY(grow-on-rawset): mirror RawSet's auto-grow.
                            // Dump emits add-part image /slideMaster[N] BEFORE the
                            // raw-set replace, so on a blank target the master slot
                            // doesn't exist yet — grow to idx so the ImagePart can
                            // attach to the right host.
                            GrowSlideMasterParts(smIdx);
                            smParts = PowerPointHandler.MastersInOrder(presentationPart);
                            if (smIdx > smParts.Count)
                                throw new ArgumentException($"slideMaster index {smIdx} out of range (total {smParts.Count})");
                        }
                        imageHost = smParts[smIdx - 1];
                    }
                    else if (slMatch != null && slMatch.Success)
                    {
                        var slIdx = int.Parse(slMatch.Groups[1].Value);
                        var slParts = PowerPointHandler.LayoutsInOrder(presentationPart);
                        if (slIdx < 1) throw new ArgumentException($"slideLayout index {slIdx} out of range");
                        if (slIdx > slParts.Count)
                        {
                            GrowSlideLayoutParts(slIdx);
                            slParts = PowerPointHandler.LayoutsInOrder(presentationPart);
                            if (slIdx > slParts.Count)
                                throw new ArgumentException($"slideLayout index {slIdx} out of range (total {slParts.Count})");
                        }
                        imageHost = slParts[slIdx - 1];
                    }
                    else if (sldMatch != null && sldMatch.Success)
                    {
                        var sldIdx = int.Parse(sldMatch.Groups[1].Value);
                        var sldParts = GetSlideParts().ToList();
                        if (sldIdx < 1 || sldIdx > sldParts.Count)
                            throw new ArgumentException($"slide index {sldIdx} out of range");
                        imageHost = sldParts[sldIdx - 1];
                    }
                    else if (nsMatch != null && nsMatch.Success)
                    {
                        var nsIdx = int.Parse(nsMatch.Groups[1].Value);
                        var sldParts = GetSlideParts().ToList();
                        if (nsIdx < 1 || nsIdx > sldParts.Count)
                            throw new ArgumentException($"noteSlide index {nsIdx} out of range");
                        // NotesSlidePart may not exist yet on the replay target.
                        // EmitNotes always lands its typed `add notes` row BEFORE
                        // this add-part, but on a blank target with no prior
                        // notes row the part is still absent; create it on
                        // demand so the ImagePart has a host to attach to.
                        // CONSISTENCY(grow-on-rawset): mirrors slideMaster/layout
                        // auto-grow above and /notesMaster on-demand creation.
                        // EnsureNotesSlidePart (not a bare AddNewPart) so the
                        // created part gets a NotesSlide root and its
                        // notesMaster relationship; the following raw-set
                        // replaces the root but keeps the relationships.
                        var hostSlide = sldParts[nsIdx - 1];
                        imageHost = EnsureNotesSlidePart(hostSlide);
                    }
                    else
                        throw new ArgumentException(
                            "add-part image: parent must be /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], or /notesMaster");
                }

                if (properties == null || !properties.TryGetValue("data", out var imgB64) || string.IsNullOrEmpty(imgB64))
                    throw new ArgumentException("add-part image requires property 'data' (base64 binary)");
                byte[] imgBytes;
                try { imgBytes = Convert.FromBase64String(imgB64); }
                catch (FormatException) { throw new ArgumentException("add-part image: 'data' is not valid base64"); }

                var imgCT = properties.TryGetValue("content-type", out var ict) && !string.IsNullOrEmpty(ict)
                    ? ict
                    : "image/png";
                var imgPartType = imgCT switch
                {
                    "image/png" => ImagePartType.Png,
                    "image/jpeg" => ImagePartType.Jpeg,
                    "image/jpg" => ImagePartType.Jpeg,
                    "image/gif" => ImagePartType.Gif,
                    "image/bmp" => ImagePartType.Bmp,
                    "image/tiff" or "image/tif" => ImagePartType.Tiff,
                    "image/x-emf" => ImagePartType.Emf,
                    "image/x-wmf" => ImagePartType.Wmf,
                    _ => ImagePartType.Png,
                };
                string? pinnedImgRid = properties.TryGetValue("rid", out var prid) && !string.IsNullOrEmpty(prid) ? prid : null;
                // rId-collision guard. The blank scaffold ships a fixed set of
                // slideLayout relationships on its master (rId1..rId5). A source
                // master whose own bg-image relationship numerically lands inside
                // that range (e.g. rId5 = master bg image, but the scaffold already
                // wired rId5 = slideLayout5) cannot pin its source rId: AddImagePart
                // with an occupied id throws, and even if it didn't, the master XML
                // raw-set's <p:bg> r:embed="rId5" would resolve to a slideLayout
                // (broken-link background → blank slide). Free the pinned id first
                // by re-homing the colliding scaffold relationship onto a fresh id.
                // For a scaffold-leftover SlideLayoutPart we additionally patch the
                // master's sldLayoutIdLst entry so the re-homed layout stays
                // referenced (mirrors GrowSlideLayoutParts' own collision fallback).
                if (!string.IsNullOrEmpty(pinnedImgRid) && imageHost is OpenXmlPartContainer collHost)
                {
                    var occupant = collHost.Parts.FirstOrDefault(p => p.RelationshipId == pinnedImgRid);
                    if (occupant.OpenXmlPart != null)
                    {
                        // Re-home the colliding scaffold relationship onto a fresh,
                        // collision-free id so the pinned id is free for the image.
                        var newRid = "Rimg" + Guid.NewGuid().ToString("N").Substring(0, 12);
                        collHost.ChangeIdOfPart(occupant.OpenXmlPart, newRid);
                        // If a master's sldLayoutIdLst still points at the old id,
                        // repoint it so the re-homed layout remains declared.
                        if (collHost is SlideMasterPart smHost && smHost.SlideMaster?.SlideLayoutIdList != null)
                        {
                            foreach (var lid in smHost.SlideMaster.SlideLayoutIdList.Elements<SlideLayoutId>())
                            {
                                if (lid.RelationshipId?.Value == pinnedImgRid)
                                {
                                    lid.RelationshipId = newRid;
                                    smHost.SlideMaster.Save();
                                    break;
                                }
                            }
                        }
                    }
                }
                ImagePart imgPart = imageHost switch
                {
                    SlidePart sp           => !string.IsNullOrEmpty(pinnedImgRid) ? sp.AddImagePart(imgPartType, pinnedImgRid) : sp.AddImagePart(imgPartType),
                    SlideMasterPart smp    => !string.IsNullOrEmpty(pinnedImgRid) ? smp.AddImagePart(imgPartType, pinnedImgRid) : smp.AddImagePart(imgPartType),
                    SlideLayoutPart sllp   => !string.IsNullOrEmpty(pinnedImgRid) ? sllp.AddImagePart(imgPartType, pinnedImgRid) : sllp.AddImagePart(imgPartType),
                    NotesMasterPart nmp    => !string.IsNullOrEmpty(pinnedImgRid) ? nmp.AddImagePart(imgPartType, pinnedImgRid) : nmp.AddImagePart(imgPartType),
                    NotesSlidePart nsp     => !string.IsNullOrEmpty(pinnedImgRid) ? nsp.AddImagePart(imgPartType, pinnedImgRid) : nsp.AddImagePart(imgPartType),
                    ThemePart tp           => !string.IsNullOrEmpty(pinnedImgRid) ? tp.AddImagePart(imgPartType, pinnedImgRid) : tp.AddImagePart(imgPartType),
                    _ => throw new ArgumentException($"add-part image: unsupported host part type {imageHost.GetType().Name}"),
                };
                using (var imgStream = new MemoryStream(imgBytes))
                    imgPart.FeedData(imgStream);
                var imgActualRid = imageHost.GetIdOfPart(imgPart);
                return (imgActualRid, parentPartPath);

            case "hyperlink":
            {
                // Re-create an EXTERNAL hyperlink relationship on a host part with
                // a pinned relationship id. Layouts/masters are emitted via raw-set
                // (wholesale XML carrying <a:hlinkClick r:id="rIdN">), but the
                // referenced external relationship is not an embedded part, so the
                // ImagePart carrier path never re-created it — the renumbered
                // rebuilt layout's .rels lost rIdN and PowerPoint rejected the file
                // ("rIdN referenced by hlinkClick does not exist"). Pinning the id
                // here makes the raw-set'd r:id="rIdN" resolve again. The host path
                // is the SOURCE-index /slideLayout[N] (or master/slide); on replay
                // GrowSlideLayoutParts maps it to the renumbered part, so the rel
                // lands on the same part the raw-set replaces.
                if (properties == null
                    || !properties.TryGetValue("target", out var hlTarget) || string.IsNullOrEmpty(hlTarget))
                    throw new ArgumentException("add-part hyperlink requires property 'target' (the external URI)");
                if (!properties.TryGetValue("rid", out var hlRid) || string.IsNullOrEmpty(hlRid))
                    throw new ArgumentException("add-part hyperlink requires property 'rid' (the relationship id to pin)");

                OpenXmlPartContainer hlHost;
                var hlSmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                var hlSlMatch = hlSmMatch.Success ? null : Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                var hlSldMatch = (hlSmMatch.Success || (hlSlMatch?.Success ?? false))
                    ? null : Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                var hlNsMatch = (hlSmMatch.Success || (hlSlMatch?.Success ?? false) || (hlSldMatch?.Success ?? false))
                    ? null : Regex.Match(parentPartPath, @"^/noteSlide\[(\d+)\]$");
                if (hlSmMatch.Success)
                {
                    var i = int.Parse(hlSmMatch.Groups[1].Value);
                    var parts = PowerPointHandler.MastersInOrder(presentationPart);
                    if (i > parts.Count) { GrowSlideMasterParts(i); parts = PowerPointHandler.MastersInOrder(presentationPart); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    hlHost = parts[i - 1];
                }
                else if (hlSlMatch != null && hlSlMatch.Success)
                {
                    var i = int.Parse(hlSlMatch.Groups[1].Value);
                    var parts = PowerPointHandler.LayoutsInOrder(presentationPart);
                    if (i > parts.Count) { GrowSlideLayoutParts(i); parts = PowerPointHandler.LayoutsInOrder(presentationPart); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideLayout index {i} out of range");
                    hlHost = parts[i - 1];
                }
                else if (hlSldMatch != null && hlSldMatch.Success)
                {
                    var i = int.Parse(hlSldMatch.Groups[1].Value);
                    var parts = GetSlideParts().ToList();
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slide index {i} out of range");
                    hlHost = parts[i - 1];
                }
                else if (hlNsMatch != null && hlNsMatch.Success)
                {
                    // CONSISTENCY(notes-image-host): mirror the add-part image
                    // /noteSlide[N] path — EmitNotes lands the typed `add notes`
                    // row first, but on a blank target with no prior notes the
                    // NotesSlidePart is still absent; create it on demand so the
                    // external hyperlink relationship has a host to attach to.
                    var i = int.Parse(hlNsMatch.Groups[1].Value);
                    var parts = GetSlideParts().ToList();
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"noteSlide index {i} out of range");
                    var hostSlide = parts[i - 1];
                    hlHost = EnsureNotesSlidePart(hostSlide);
                }
                else
                    throw new ArgumentException(
                        "add-part hyperlink: parent must be /slideLayout[N], /slideMaster[N], /slide[N], or /noteSlide[N]");

                // Idempotent: if a relationship with this id already exists, don't
                // re-add (AddHyperlinkRelationship throws on a duplicate id).
                if (hlHost.HyperlinkRelationships.Any(r => r.Id == hlRid))
                    return (hlRid, parentPartPath);
                hlHost.AddHyperlinkRelationship(new Uri(hlTarget, UriKind.RelativeOrAbsolute), isExternal: true, hlRid);
                return (hlRid, parentPartPath);
            }

            case "tags":
            {
                // Re-create a UserDefinedTags part (programmability metadata,
                // <p:tagLst>) on a host part with a pinned relationship id and the
                // source tag XML. Layouts/masters are emitted via raw-set (wholesale
                // XML carrying <p:custDataLst><p:tags r:id="rIdN"/>), but the tags
                // part lives in the host's own .rels enumerated separately, so the
                // ImagePart/hyperlink carriers never re-created it — the renumbered
                // rebuilt layout's r:id="rIdN" dangled and PowerPoint rejected the
                // whole deck (0x80070570 OPC corrupt). Pinning the id here makes the
                // raw-set'd reference resolve. The host path is the SOURCE-index
                // /slideLayout[N] (or master); on replay GrowSlideLayoutParts maps
                // it to the renumbered part, so the tags rel lands on the same part
                // the raw-set replaces. (mirrors the add-part hyperlink pattern)
                if (properties == null
                    || !properties.TryGetValue("data", out var tagXml) || string.IsNullOrEmpty(tagXml))
                    throw new ArgumentException("add-part tags requires property 'data' (the <p:tagLst> XML)");
                if (!properties.TryGetValue("rid", out var tagRid) || string.IsNullOrEmpty(tagRid))
                    throw new ArgumentException("add-part tags requires property 'rid' (the relationship id to pin)");

                OpenXmlPartContainer tagHost;
                var tgSmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                var tgSlMatch = tgSmMatch.Success ? null : Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                var tgSldMatch = (tgSmMatch.Success || (tgSlMatch?.Success ?? false))
                    ? null : Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (tgSmMatch.Success)
                {
                    var i = int.Parse(tgSmMatch.Groups[1].Value);
                    var parts = PowerPointHandler.MastersInOrder(presentationPart);
                    if (i > parts.Count) { GrowSlideMasterParts(i); parts = PowerPointHandler.MastersInOrder(presentationPart); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    tagHost = parts[i - 1];
                }
                else if (tgSlMatch != null && tgSlMatch.Success)
                {
                    var i = int.Parse(tgSlMatch.Groups[1].Value);
                    var parts = PowerPointHandler.LayoutsInOrder(presentationPart);
                    if (i > parts.Count) { GrowSlideLayoutParts(i); parts = PowerPointHandler.LayoutsInOrder(presentationPart); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideLayout index {i} out of range");
                    tagHost = parts[i - 1];
                }
                else if (tgSldMatch != null && tgSldMatch.Success)
                {
                    var i = int.Parse(tgSldMatch.Groups[1].Value);
                    var parts = GetSlideParts().ToList();
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slide index {i} out of range");
                    tagHost = parts[i - 1];
                }
                else
                    throw new ArgumentException(
                        "add-part tags: parent must be /slideLayout[N], /slideMaster[N], or /slide[N]");

                // The blank scaffold's master ships rId1..rId5 as slideLayout
                // relationships, so a master/layout <p:tags r:id="rId5"> collides
                // with a scaffold layout. Re-home the colliding part onto a fresh
                // id (repointing the master's sldLayoutIdLst) so the pinned tag id
                // is free — otherwise the tag part would silently not be created
                // and the raw-set'd r:id dangled. Mirrors the add-part image /
                // extpart collision path. (A genuine same-rId re-run is a no-op
                // since AddNewPart would then create a duplicate — but per batch
                // each tag rId is emitted once.)
                if (tagHost is OpenXmlPartContainer tagColl)
                {
                    var occ = tagColl.Parts.FirstOrDefault(p => p.RelationshipId == tagRid);
                    if (occ.OpenXmlPart is UserDefinedTagsPart)
                        return (tagRid, parentPartPath); // already the tag part — idempotent
                    ReHomeCollidingRel(tagColl, tagRid);
                }
                var newTagPart = tagHost switch
                {
                    SlidePart sp        => sp.AddNewPart<UserDefinedTagsPart>(tagRid),
                    SlideLayoutPart slp => slp.AddNewPart<UserDefinedTagsPart>(tagRid),
                    SlideMasterPart smp => smp.AddNewPart<UserDefinedTagsPart>(tagRid),
                    _ => throw new ArgumentException($"add-part tags: unsupported host part type {tagHost.GetType().Name}"),
                };
                using (var sw = new StreamWriter(newTagPart.GetStream(FileMode.Create, FileAccess.Write), new System.Text.UTF8Encoding(false)))
                    sw.Write(tagXml);
                return (tagHost.GetIdOfPart(newTagPart), parentPartPath);
            }

            case "sliderel":
            {
                // Pin an internal slide-jump relationship (type .../slide) so a
                // raw-carried <a:hlinkClick r:id="rIdN" action="…hlinksldjump">
                // (e.g. inside a table cell's txBodyRaw) resolves to the rebuilt
                // target slide. Must replay AFTER every slide exists — the
                // emitter defers it. Props: rid (pinned), target (1-based ordinal
                // of the target slide).
                if (properties == null
                    || !properties.TryGetValue("rid", out var srRid) || string.IsNullOrEmpty(srRid))
                    throw new ArgumentException("add-part sliderel requires property 'rid'");
                if (!properties.TryGetValue("target", out var srTgt)
                    || !int.TryParse(srTgt, out var srTgtOrd))
                    throw new ArgumentException("add-part sliderel requires property 'target' (1-based slide ordinal)");
                var srMatch = Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (!srMatch.Success)
                    throw new ArgumentException("add-part sliderel: parent must be /slide[N]");
                var srParts = GetSlideParts().ToList();
                var srHostIdx = int.Parse(srMatch.Groups[1].Value);
                if (srHostIdx < 1 || srHostIdx > srParts.Count)
                    throw new ArgumentException($"slide index {srHostIdx} out of range");
                if (srTgtOrd < 1 || srTgtOrd > srParts.Count)
                    throw new ArgumentException($"sliderel target ordinal {srTgtOrd} out of range (total {srParts.Count})");
                var srHost = srParts[srHostIdx - 1];
                // Idempotent: skip if the rId is already wired.
                if (srHost.Parts.Any(p => p.RelationshipId == srRid)
                    || srHost.ExternalRelationships.Any(r => r.Id == srRid))
                    return (srRid, parentPartPath);
                srHost.AddPart(srParts[srTgtOrd - 1], srRid);
                return (srRid, parentPartPath);
            }

            case "tablestyles":
            {
                // Re-create the presentation's custom table-style catalogue
                // (ppt/tableStyles.xml). MUST use the TYPED TableStylesPart, not
                // AddExtendedPart: tableStyles is a known OOXML part type, and a
                // generic extended part with that content-type is pruned by the
                // SDK on save (the part vanished and each table's custom
                // <a:tableStyleId> GUID fell back to a built-in style —
                // sample09). Props: xml (base64 of tableStyles.xml).
                var tsPres = _doc.PresentationPart
                    ?? throw new InvalidOperationException("presentation part missing");
                if (properties == null
                    || !properties.TryGetValue("xml", out var tsXmlB64) || string.IsNullOrEmpty(tsXmlB64))
                    throw new ArgumentException("add-part tablestyles requires property 'xml' (base64)");
                byte[] tsBytes;
                try { tsBytes = Convert.FromBase64String(tsXmlB64); }
                catch (FormatException) { throw new ArgumentException("add-part tablestyles: 'xml' is not valid base64"); }
                var tsPart = tsPres.TableStylesPart ?? tsPres.AddNewPart<TableStylesPart>();
                using (var tsStream = tsPart.GetStream(FileMode.Create, FileAccess.Write))
                    tsStream.Write(tsBytes, 0, tsBytes.Length);
                return ("tablestyles", parentPartPath);
            }

            case "extpart":
            {
                // Re-create an arbitrary binary part with a CUSTOM relationship
                // type + pinned relationship id. Used to carry a picture's blip
                // companion parts — the HD Photo backup layer (.wdp, rel type
                // .../hdphoto) and SVG companion — that the typed `add picture`
                // path doesn't reproduce. The blip's extLst is re-appended
                // verbatim (passthrough), keeping <... r:embed="rIdN">, so the
                // companion part must exist with the SAME rId or the rebuilt
                // picture carries a dangling relationship (lost effects layer;
                // strict consumers reject the package). Mirrors add-part image
                // but preserves the source relationship type via AddExtendedPart.
                if (properties == null
                    || !properties.TryGetValue("data", out var epB64) || string.IsNullOrEmpty(epB64))
                    throw new ArgumentException("add-part extpart requires property 'data' (base64 binary)");
                if (!properties.TryGetValue("rid", out var epRid) || string.IsNullOrEmpty(epRid))
                    throw new ArgumentException("add-part extpart requires property 'rid'");
                if (!properties.TryGetValue("rel-type", out var epRelType) || string.IsNullOrEmpty(epRelType))
                    throw new ArgumentException("add-part extpart requires property 'rel-type' (the relationship type URI)");
                var epContentType = properties.TryGetValue("content-type", out var epct) && !string.IsNullOrEmpty(epct)
                    ? epct : "application/octet-stream";
                var epExt = properties.TryGetValue("ext", out var epe) && !string.IsNullOrEmpty(epe) ? epe : ".bin";
                byte[] epBytes;
                try { epBytes = Convert.FromBase64String(epB64); }
                catch (FormatException) { throw new ArgumentException("add-part extpart: 'data' is not valid base64"); }

                OpenXmlPartContainer epHost;
                // Presentation-level custom binary part (e.g. Google Slides'
                // ppt/metadata, reached by <go:slidesCustomData r:id="rIdN"> inside
                // the presentation extLst). The extLst is replayed via raw-set, so
                // the part + relationship must be re-pinned on the presentation
                // part or the r:id dangles and PowerPoint refuses the deck.
                if (parentPartPath == "/presentation")
                {
                    epHost = presentationPart;
                    // A non-part (external/hyperlink) relationship can't be
                    // re-homed by ChangeIdOfPart, and silently skipping here
                    // would report success while dropping the binary part.
                    // Fail loudly instead — under atomic batch the file is
                    // rolled back untouched. Unreachable from a self-produced
                    // dump (both carriers pin ids from the same source rels).
                    if (epHost.ExternalRelationships.Any(r => r.Id == epRid)
                        || epHost.HyperlinkRelationships.Any(r => r.Id == epRid))
                        throw new ArgumentException(
                            $"add-part extpart: rid '{epRid}' already exists as an external/hyperlink relationship on /presentation and cannot be re-homed. Use a different rid.");
                    // Part-rel collision: on a rebuilt deck the scaffold's own
                    // rels (master/slide/presProps…) occupy low rIds, so a
                    // pinned source id like rId2 is usually TAKEN. Skipping
                    // here (the old behavior) silently dropped the part while
                    // still reporting success. Idempotent-skip only a genuine
                    // re-run (same-rel-type ExtendedPart already present);
                    // otherwise re-home the occupant so the pinned id is free.
                    // Mirrors the slide/master/layout branch below.
                    var epOcc = epHost.Parts.FirstOrDefault(p => p.RelationshipId == epRid);
                    if (epOcc.OpenXmlPart is ExtendedPart occExt && occExt.RelationshipType == epRelType)
                        return (epRid, parentPartPath);
                    ReHomeCollidingRel(epHost, epRid);
                    var epPresPart = epHost.AddExtendedPart(epRelType, epContentType, epExt, epRid);
                    using (var epStream = new MemoryStream(epBytes))
                        epPresPart.FeedData(epStream);
                    return (epRid, parentPartPath);
                }
                var epSmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                var epSlMatch = epSmMatch.Success ? null : Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                var epSldMatch = (epSmMatch.Success || (epSlMatch?.Success ?? false))
                    ? null : Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (epSmMatch.Success)
                {
                    var i = int.Parse(epSmMatch.Groups[1].Value);
                    var ps = PowerPointHandler.MastersInOrder(presentationPart);
                    if (i > ps.Count) { GrowSlideMasterParts(i); ps = PowerPointHandler.MastersInOrder(presentationPart); }
                    if (i < 1 || i > ps.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    epHost = ps[i - 1];
                }
                else if (epSlMatch != null && epSlMatch.Success)
                {
                    var i = int.Parse(epSlMatch.Groups[1].Value);
                    var ps = PowerPointHandler.LayoutsInOrder(presentationPart);
                    if (i > ps.Count) { GrowSlideLayoutParts(i); ps = PowerPointHandler.LayoutsInOrder(presentationPart); }
                    if (i < 1 || i > ps.Count) throw new ArgumentException($"slideLayout index {i} out of range");
                    epHost = ps[i - 1];
                }
                else if (epSldMatch != null && epSldMatch.Success)
                {
                    var i = int.Parse(epSldMatch.Groups[1].Value);
                    var ps = GetSlideParts().ToList();
                    if (i < 1 || i > ps.Count) throw new ArgumentException($"slide index {i} out of range");
                    epHost = ps[i - 1];
                }
                else
                    throw new ArgumentException(
                        "add-part extpart: parent must be /presentation, /slide[N], /slideLayout[N], or /slideMaster[N]");

                // External/hyperlink-rel collision: a non-part relationship
                // can't be re-homed by ChangeIdOfPart, and the old idempotent
                // skip reported success while dropping the binary part. Fail
                // loudly instead (atomic batch rolls the file back).
                // Unreachable from a self-produced dump — both carriers pin
                // ids from the same source rels, so they never collide.
                // Part collision (scaffold layout rel occupying rId3..rId5 on
                // the master): re-home it so the pinned id is free — otherwise
                // the extpart silently skips and the hdphoto r:embed dangles.
                // Mirrors the add-part image collision path.
                if (epHost.ExternalRelationships.Any(r => r.Id == epRid)
                    || epHost.HyperlinkRelationships.Any(r => r.Id == epRid))
                    throw new ArgumentException(
                        $"add-part extpart: rid '{epRid}' already exists as an external/hyperlink relationship on {parentPartPath} and cannot be re-homed. Use a different rid.");
                // Genuine re-run (same-rel-type ExtendedPart already pinned on
                // this rid): idempotent skip, same as the /presentation branch —
                // re-homing our own part would duplicate it under a fresh id.
                var epSlideOcc = epHost.Parts.FirstOrDefault(p => p.RelationshipId == epRid);
                if (epSlideOcc.OpenXmlPart is ExtendedPart epSlideExt && epSlideExt.RelationshipType == epRelType)
                    return (epRid, parentPartPath);
                ReHomeCollidingRel(epHost, epRid);
                var epPart = epHost.AddExtendedPart(epRelType, epContentType, epExt, epRid);
                using (var epStream = new MemoryStream(epBytes))
                    epPart.FeedData(epStream);
                return (epRid, parentPartPath);
            }

            case "activex":
            {
                // Carry an ActiveX control part for round-trip: activeX{N}.xml
                // (rel type .../control on the slide, pinned rId matching the
                // <p:control r:id>) + its child activeX{N}.bin (the control's own
                // .rels, activeXControlBinary rel, pinned rId matching the
                // activeX xml's internal r:id). Without the parts the rIds in the
                // round-tripped <p:controls> block dangle → 0x80070570.
                var axm = System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (!axm.Success)
                    throw new ArgumentException("add-part activex: parent must be /slide[N]");
                var axIdx = int.Parse(axm.Groups[1].Value);
                var axParts = GetSlideParts().ToList();
                if (axIdx < 1 || axIdx > axParts.Count)
                    throw new ArgumentException($"slide index {axIdx} out of range");
                var axSlide = axParts[axIdx - 1];

                if (properties == null
                    || !properties.TryGetValue("rid", out var axRid) || string.IsNullOrEmpty(axRid))
                    throw new ArgumentException("add-part activex requires property 'rid'");
                if (!properties.TryGetValue("xml", out var axXmlB64) || string.IsNullOrEmpty(axXmlB64))
                    throw new ArgumentException("add-part activex requires property 'xml' (base64)");
                byte[] axXmlBytes;
                try { axXmlBytes = Convert.FromBase64String(axXmlB64); }
                catch (FormatException) { throw new ArgumentException("add-part activex: 'xml' is not valid base64"); }

                const string ControlRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
                const string ActiveXCT = "application/vnd.ms-office.activeX+xml";
                const string BinRelType = "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
                const string BinCT = "application/vnd.ms-office.activeX";

                if (axSlide.ExternalRelationships.Any(r => r.Id == axRid)
                    || axSlide.HyperlinkRelationships.Any(r => r.Id == axRid))
                    return (axRid, parentPartPath);
                ReHomeCollidingRel(axSlide, axRid);
                var axPart = axSlide.AddExtendedPart(ControlRelType, ActiveXCT, ".xml", axRid);
                using (var axStream = new MemoryStream(axXmlBytes))
                    axPart.FeedData(axStream);

                if (properties.TryGetValue("bin", out var axBinB64) && !string.IsNullOrEmpty(axBinB64)
                    && properties.TryGetValue("bin-rid", out var axBinRid) && !string.IsNullOrEmpty(axBinRid))
                {
                    byte[] axBinBytes;
                    try { axBinBytes = Convert.FromBase64String(axBinB64); }
                    catch (FormatException) { axBinBytes = Array.Empty<byte>(); }
                    if (axBinBytes.Length > 0)
                    {
                        var binPart = axPart.AddExtendedPart(BinRelType, BinCT, ".bin", axBinRid);
                        using var binStream = new MemoryStream(axBinBytes);
                        binPart.FeedData(binStream);
                    }
                }
                return (axRid, parentPartPath);
            }

            case "chartex":
            {
                // Carry a chartEx part (cx: extension charts — funnel, sunburst,
                // treemap, …) for round-trip: chartEx{N}.xml with a pinned rId
                // matching the AlternateContent slice's <cx:chart r:id>, plus its
                // typed children (colors / style sidecars and the embedded xlsx
                // workbook) with pinned rIds matching the chartEx XML's internal
                // references. Without the parts the slice's rIds dangle and
                // PowerPoint refuses the deck (0x80070570, funnel-pp1). Typed
                // ExtendedChartPart is REQUIRED — a generic AddExtendedPart of a
                // known part type gets pruned on save (tableStyles lesson).
                var cxm = System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (!cxm.Success)
                    throw new ArgumentException("add-part chartex: parent must be /slide[N]");
                var cxIdx = int.Parse(cxm.Groups[1].Value);
                var cxParts = GetSlideParts().ToList();
                if (cxIdx < 1 || cxIdx > cxParts.Count)
                    throw new ArgumentException($"slide index {cxIdx} out of range");
                var cxSlide = cxParts[cxIdx - 1];

                if (properties == null
                    || !properties.TryGetValue("rid", out var cxRid) || string.IsNullOrEmpty(cxRid))
                    throw new ArgumentException("add-part chartex requires property 'rid'");
                if (!properties.TryGetValue("xml", out var cxXmlB64) || string.IsNullOrEmpty(cxXmlB64))
                    throw new ArgumentException("add-part chartex requires property 'xml' (base64)");
                byte[] cxXmlBytes;
                try { cxXmlBytes = Convert.FromBase64String(cxXmlB64); }
                catch (FormatException) { throw new ArgumentException("add-part chartex: 'xml' is not valid base64"); }

                if (cxSlide.Parts.Any(p => p.RelationshipId == cxRid))
                    return (cxRid, parentPartPath);
                ReHomeCollidingRel(cxSlide, cxRid);
                var cxPart = cxSlide.AddNewPart<ExtendedChartPart>(cxRid);
                using (var cxs = new MemoryStream(cxXmlBytes)) cxPart.FeedData(cxs);

                void FeedChild<T>(string ridKey, string dataKey) where T : OpenXmlPart, IFixedContentTypePart
                {
                    if (!properties.TryGetValue(ridKey, out var crid) || string.IsNullOrEmpty(crid)) return;
                    if (!properties.TryGetValue(dataKey, out var cb64) || string.IsNullOrEmpty(cb64)) return;
                    byte[] cb;
                    try { cb = Convert.FromBase64String(cb64); } catch (FormatException) { return; }
                    var child = cxPart.AddNewPart<T>(crid);
                    using var cs = new MemoryStream(cb);
                    child.FeedData(cs);
                }
                FeedChild<ChartColorStylePart>("colors-rid", "colors");
                FeedChild<ChartStylePart>("style-rid", "style");
                if (properties.TryGetValue("package-rid", out var pkgRid) && !string.IsNullOrEmpty(pkgRid)
                    && properties.TryGetValue("package", out var pkgB64) && !string.IsNullOrEmpty(pkgB64))
                {
                    byte[] pkgBytes;
                    try { pkgBytes = Convert.FromBase64String(pkgB64); } catch (FormatException) { pkgBytes = Array.Empty<byte>(); }
                    if (pkgBytes.Length > 0)
                    {
                        var pkgCT = properties.GetValueOrDefault("package-content-type",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                        var pkgPart = cxPart.AddNewPart<EmbeddedPackagePart>(pkgCT, pkgRid);
                        using var ps = new MemoryStream(pkgBytes);
                        pkgPart.FeedData(ps);
                    }
                }
                return (cxRid, parentPartPath);
            }

            case "chartimage":
            {
                // Attach an ImagePart to a slide's Nth ChartPart with a pinned
                // rId — re-creates the image a chart-internal <a:blipFill
                // r:embed> references (chart/plot-area picture or texture
                // fill). Without it the verbatim spPr fill dangles and the
                // emitter had to degrade it to noFill (chart-picture-bg).
                var cim = System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (!cim.Success)
                    throw new ArgumentException("add-part chartimage: parent must be /slide[N]");
                if (properties == null
                    || !properties.TryGetValue("rid", out var ciRid) || string.IsNullOrEmpty(ciRid))
                    throw new ArgumentException("add-part chartimage requires property 'rid'");
                if (!properties.TryGetValue("chart", out var ciOrdRaw) || !int.TryParse(ciOrdRaw, out var ciOrd))
                    throw new ArgumentException("add-part chartimage requires property 'chart' (1-based chart ordinal)");
                if (!properties.TryGetValue("data", out var ciB64) || string.IsNullOrEmpty(ciB64))
                    throw new ArgumentException("add-part chartimage requires property 'data' (base64)");
                var ciCT = properties.GetValueOrDefault("content-type", "image/png");
                var ciIdx = int.Parse(cim.Groups[1].Value);
                var ciSlides = GetSlideParts().ToList();
                if (ciIdx < 1 || ciIdx > ciSlides.Count)
                    throw new ArgumentException($"slide index {ciIdx} out of range");
                var ciCharts = ciSlides[ciIdx - 1].ChartParts.ToList();
                if (ciOrd < 1 || ciOrd > ciCharts.Count)
                    throw new ArgumentException($"chart ordinal {ciOrd} out of range (total: {ciCharts.Count})");
                var ciChart = ciCharts[ciOrd - 1];
                if (ciChart.Parts.Any(pp2 => pp2.RelationshipId == ciRid))
                    return (ciRid, parentPartPath);
                var ciPart = ciChart.AddImagePart(ciCT, ciRid);
                byte[] ciBytes;
                try { ciBytes = Convert.FromBase64String(ciB64); }
                catch (FormatException) { throw new ArgumentException("add-part chartimage: 'data' is not valid base64"); }
                using (var cis = new MemoryStream(ciBytes)) ciPart.FeedData(cis);
                return (ciRid, parentPartPath);
            }

            case "chartstyle":
            {
                // Attach a ChartStylePart / ChartColorStylePart to a slide's
                // Nth ChartPart with a pinned rId (counterpart of
                // GetChartStyleParts). Props: chart, kind=style|colors,
                // rid, data (base64 XML).
                var csm = System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (!csm.Success)
                    throw new ArgumentException("add-part chartstyle: parent must be /slide[N]");
                if (properties == null
                    || !properties.TryGetValue("rid", out var csRid) || string.IsNullOrEmpty(csRid))
                    throw new ArgumentException("add-part chartstyle requires property 'rid'");
                if (!properties.TryGetValue("chart", out var csOrdRaw) || !int.TryParse(csOrdRaw, out var csOrd))
                    throw new ArgumentException("add-part chartstyle requires property 'chart' (1-based chart ordinal)");
                if (!properties.TryGetValue("kind", out var csKind)
                    || csKind is not ("style" or "colors" or "themeoverride"))
                    throw new ArgumentException("add-part chartstyle requires property 'kind' (style|colors|themeoverride)");
                if (!properties.TryGetValue("data", out var csB64) || string.IsNullOrEmpty(csB64))
                    throw new ArgumentException("add-part chartstyle requires property 'data' (base64)");
                var csIdx = int.Parse(csm.Groups[1].Value);
                var csSlides = GetSlideParts().ToList();
                if (csIdx < 1 || csIdx > csSlides.Count)
                    throw new ArgumentException($"slide index {csIdx} out of range");
                var csCharts = csSlides[csIdx - 1].ChartParts.ToList();
                if (csOrd < 1 || csOrd > csCharts.Count)
                    throw new ArgumentException($"chart ordinal {csOrd} out of range (total: {csCharts.Count})");
                var csChart = csCharts[csOrd - 1];
                if (csChart.Parts.Any(pp3 => pp3.RelationshipId == csRid))
                    return (csRid, parentPartPath);
                byte[] csBytes;
                try { csBytes = Convert.FromBase64String(csB64); }
                catch (FormatException) { throw new ArgumentException("add-part chartstyle: 'data' is not valid base64"); }
                OpenXmlPart csPart = csKind switch
                {
                    "style" => csChart.AddNewPart<ChartStylePart>(csRid),
                    "colors" => csChart.AddNewPart<ChartColorStylePart>(csRid),
                    _ => csChart.AddNewPart<ThemeOverridePart>(csRid),
                };
                using (var css = new MemoryStream(csBytes)) csPart.FeedData(css);
                return (csRid, parentPartPath);
            }

            case "extrel":
            {
                // Re-create an EXTERNAL relationship (TargetMode=External) with a
                // pinned id and a specified relationship type — used to carry a
                // master/layout picture's external image link (<a:blip r:link>),
                // which the embedded-ImagePart carrier doesn't cover. Props:
                // rid, rel-type, target (the external URI).
                if (properties == null
                    || !properties.TryGetValue("rid", out var erRid) || string.IsNullOrEmpty(erRid))
                    throw new ArgumentException("add-part extrel requires property 'rid'");
                if (!properties.TryGetValue("rel-type", out var erType) || string.IsNullOrEmpty(erType))
                    throw new ArgumentException("add-part extrel requires property 'rel-type'");
                if (!properties.TryGetValue("target", out var erTarget) || string.IsNullOrEmpty(erTarget))
                    throw new ArgumentException("add-part extrel requires property 'target'");
                OpenXmlPartContainer erHost;
                var erSm = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                var erSl = erSm.Success ? null : Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                var erSld = (erSm.Success || (erSl?.Success ?? false)) ? null : Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (erSm.Success)
                {
                    var i = int.Parse(erSm.Groups[1].Value);
                    var ps = PowerPointHandler.MastersInOrder(presentationPart);
                    if (i > ps.Count) { GrowSlideMasterParts(i); ps = PowerPointHandler.MastersInOrder(presentationPart); }
                    if (i < 1 || i > ps.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    erHost = ps[i - 1];
                }
                else if (erSl != null && erSl.Success)
                {
                    var i = int.Parse(erSl.Groups[1].Value);
                    var ps = PowerPointHandler.LayoutsInOrder(presentationPart);
                    if (i > ps.Count) { GrowSlideLayoutParts(i); ps = PowerPointHandler.LayoutsInOrder(presentationPart); }
                    if (i < 1 || i > ps.Count) throw new ArgumentException($"slideLayout index {i} out of range");
                    erHost = ps[i - 1];
                }
                else if (erSld != null && erSld.Success)
                {
                    var i = int.Parse(erSld.Groups[1].Value);
                    var ps = GetSlideParts().ToList();
                    if (i < 1 || i > ps.Count) throw new ArgumentException($"slide index {i} out of range");
                    erHost = ps[i - 1];
                }
                else
                    throw new ArgumentException("add-part extrel: parent must be /slide[N], /slideMaster[N] or /slideLayout[N]");
                // Idempotent.
                if (erHost.ExternalRelationships.Any(r => r.Id == erRid)
                    || erHost.Parts.Any(p => p.RelationshipId == erRid))
                    return (erRid, parentPartPath);
                erHost.AddExternalRelationship(erType, new Uri(erTarget, UriKind.RelativeOrAbsolute), erRid);
                return (erRid, parentPartPath);
            }

            case "theme":
            {
                // Attach a DISTINCT theme part to a slideMaster / notesMaster with
                // a pinned relationship id and the source theme XML. Multi-master
                // decks give each master its own theme; the blank scaffold +
                // GrowSlideMasterParts share the presentation's primary theme,
                // which loses theme2/theme3 content and (worse) makes masters
                // reference the wrong / a shared theme — PowerPoint refuses such a
                // deck. This re-creates each master's own theme so the package
                // matches the source's 1:1 master:theme topology.
                if (properties == null
                    || !properties.TryGetValue("data", out var themeXml) || string.IsNullOrEmpty(themeXml))
                    throw new ArgumentException("add-part theme requires property 'data' (the theme XML)");
                // The theme is a part relationship of the master, NOT referenced
                // by r:id anywhere in the master XML body, so the relationship id
                // need not be pinned — and pinning it risks colliding with the
                // master's other relationships (images/layouts). Let the SDK
                // assign a fresh id; only honour an explicit rid when it's free.
                string? themeRid = properties.TryGetValue("rid", out var trid) && !string.IsNullOrEmpty(trid) ? trid : null;

                OpenXmlPartContainer themeHost;
                var tmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                if (tmMatch.Success)
                {
                    var i = int.Parse(tmMatch.Groups[1].Value);
                    var parts = PowerPointHandler.MastersInOrder(presentationPart);
                    if (i > parts.Count) { GrowSlideMasterParts(i); parts = PowerPointHandler.MastersInOrder(presentationPart); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    themeHost = parts[i - 1];
                }
                else if (parentPartPath == "/notesMaster")
                {
                    themeHost = presentationPart.NotesMasterPart
                        ?? throw new ArgumentException("add-part theme: notesMaster part does not exist yet");
                }
                else
                    throw new ArgumentException(
                        "add-part theme: parent must be /slideMaster[N] or /notesMaster");

                // Remove any existing (shared/placeholder) theme part on this host
                // so it gets its OWN distinct part rather than pointing at the
                // primary theme. Then add a fresh ThemePart with the pinned rId.
                var existingTheme = themeHost switch
                {
                    SlideMasterPart smp => (ThemePart?)smp.ThemePart,
                    NotesMasterPart nmp => nmp.ThemePart,
                    _ => null,
                };
                if (existingTheme != null)
                    themeHost.DeletePart(existingTheme);

                // Only pin the rId if it's not already taken by another rel on the
                // host (after deleting the old theme). Otherwise auto-assign.
                bool ridFree = !string.IsNullOrEmpty(themeRid)
                    && !themeHost.Parts.Any(p => p.RelationshipId == themeRid)
                    && !themeHost.ExternalRelationships.Any(r => r.Id == themeRid)
                    && !themeHost.HyperlinkRelationships.Any(r => r.Id == themeRid);
                ThemePart newTheme = themeHost switch
                {
                    SlideMasterPart smp => ridFree ? smp.AddNewPart<ThemePart>(themeRid!) : smp.AddNewPart<ThemePart>(),
                    NotesMasterPart nmp => ridFree ? nmp.AddNewPart<ThemePart>(themeRid!) : nmp.AddNewPart<ThemePart>(),
                    _ => throw new ArgumentException($"add-part theme: unsupported host {themeHost.GetType().Name}"),
                };
                using (var ts = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(themeXml)))
                    newTheme.FeedData(ts);
                // Re-attach texture images the theme's fmtScheme references via
                // <a:blipFill r:embed="rIdN">; without them the verbatim theme XML
                // dangles. Pinned rIds match the fed theme XML. (Same carrier as
                // diagram/picture images — flat numbered themeImage{k} props.)
                AttachDiagramImages(newTheme, properties, "themeImage");
                var themeActualRid = themeHost.GetIdOfPart(newTheme);
                return (themeActualRid, parentPartPath);
            }

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, smartart, video, audio, model3d, ole, image, hyperlink, theme");
        }
    }

    // Write verbatim XML into a freshly-created SmartArt diagram sub-part, or
    // seed a minimal typed root when no content was supplied. Writing the raw
    // bytes directly (not via the typed root) preserves the source's exact
    // namespace declarations / extension prefixes that the dump emitter
    // canonicalised, and — crucially — lands the content regardless of the
    // SDK's part-naming base (the parts land under /ppt/graphics/, not the
    // source's /ppt/diagrams/, so a URI-targeted raw-set could not reach them).
    // Free a pinned relationship id on a host by re-homing whatever part
    // currently occupies it onto a fresh id. The blank scaffold ships a master
    // with rId1..rId5 = slideLayouts, so pinning a source image / extended-part
    // rId that lands in that range collides; without re-homing, an add-part that
    // idempotent-skips on collision silently fails to create the part and its
    // r:embed dangles. For a scaffold SlideLayoutPart occupant we also repoint
    // the master's sldLayoutIdLst entry so the re-homed layout stays declared.
    // Mirrors the inline re-home in the add-part image case.
    private static void ReHomeCollidingRel(OpenXmlPartContainer host, string pinnedRid)
    {
        if (string.IsNullOrEmpty(pinnedRid)) return;
        var occupant = host.Parts.FirstOrDefault(p => p.RelationshipId == pinnedRid);
        if (occupant.OpenXmlPart == null) return;
        var newRid = "Rreh" + Guid.NewGuid().ToString("N").Substring(0, 12);
        host.ChangeIdOfPart(occupant.OpenXmlPart, newRid);
        if (host is SlideMasterPart smHost && smHost.SlideMaster?.SlideLayoutIdList != null)
        {
            foreach (var lid in smHost.SlideMaster.SlideLayoutIdList.Elements<SlideLayoutId>())
            {
                if (lid.RelationshipId?.Value == pinnedRid)
                {
                    lid.RelationshipId = newRid;
                    smHost.SlideMaster.Save();
                    break;
                }
            }
        }
        // Presentation-part host: every r:id inside presentation.xml resolves
        // against the presentation part's rels (sldMasterIdLst / sldIdLst /
        // notesMasterIdLst / custShow slide lists), so repoint any attribute
        // still carrying the vacated id or the deck dangles on open.
        else if (host is PresentationPart ppHost && ppHost.Presentation != null)
        {
            const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var changed = false;
            foreach (var el in ppHost.Presentation.Descendants())
            {
                foreach (var attr in el.GetAttributes())
                {
                    if (attr.LocalName == "id" && attr.NamespaceUri == RelNs && attr.Value == pinnedRid)
                    {
                        el.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, newRid));
                        changed = true;
                    }
                }
            }
            if (changed) ppHost.Presentation.Save();
        }
    }

    /// <summary>
    /// Create a slide sub-part with a pinned relationship id, OR reuse the
    /// existing part when that rId is already present. Two SmartArt
    /// graphicFrames on one slide may SHARE a diagram layout/colors/quickStyle
    /// part (distinct data models, common styling — e.g. data1 + data2 both
    /// referencing a single layout1/colors1/quickStyle1 set).
    /// The dump emitter emits the shared rId on BOTH add-part ops, so the second
    /// call must reuse the first call's part instead of re-pinning the same rId,
    /// which otherwise throws "Id conflicts with the Id of an existing
    /// relationship" and leaves the package unopenable (0x80070570).
    /// <paramref name="created"/> is false when an existing part was reused, so
    /// the caller can skip rewriting content already written by the first call.
    /// </summary>
    private static T GetOrAddPinnedPart<T>(SlidePart slidePart, string? pinnedRid, out bool created)
        where T : OpenXmlPart, DocumentFormat.OpenXml.Packaging.IFixedContentTypePart
    {
        if (!string.IsNullOrEmpty(pinnedRid))
        {
            var existing = slidePart.Parts.FirstOrDefault(p => p.RelationshipId == pinnedRid).OpenXmlPart;
            if (existing is T typed) { created = false; return typed; }
            // rId taken by a different part type (malformed dump): let the SDK
            // allocate a fresh rId rather than collide.
            if (existing != null) { created = true; return slidePart.AddNewPart<T>(); }
            created = true;
            return slidePart.AddNewPart<T>(pinnedRid);
        }
        created = true;
        return slidePart.AddNewPart<T>();
    }

    private static void WriteDiagramPartXml(
        OpenXmlPart part, string? xml, Func<OpenXmlElement> seedFactory)
    {
        if (!string.IsNullOrEmpty(xml))
        {
            const string prolog = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
            writer.Write(prolog);
            writer.Write(xml);
            return;
        }
        // Fallback: minimal typed root so direct CLI add-part stays usable.
        var seed = seedFactory();
        using var seedStream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var xw = System.Xml.XmlWriter.Create(seedStream,
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = false, Encoding = new System.Text.UTF8Encoding(false) });
        seed.WriteTo(xw);
    }

    // Re-attach the pictures a diagram data / drawing part references via its own
    // .rels. The emitter flattens GetSmartArtsOnSlide's DataImages/DrawingImages
    // into numbered props ({prefix}{k}.rid/.ct/.data); replay recreates each
    // ImagePart on the host with the SOURCE rId pinned, so the part XML's
    // <a:blip r:embed="rIdN"> resolves instead of dangling (which otherwise makes
    // PowerPoint refuse the deck). No-op when no props with the prefix are present.
    private static void AttachDiagramImages(OpenXmlPart host, Dictionary<string, string>? properties, string prefix)
    {
        if (properties == null) return;
        for (int k = 0; ; k++)
        {
            if (!properties.TryGetValue($"{prefix}{k}.rid", out var rid) || string.IsNullOrEmpty(rid))
                break;
            properties.TryGetValue($"{prefix}{k}.ct", out var ct);
            properties.TryGetValue($"{prefix}{k}.data", out var b64);
            if (string.IsNullOrEmpty(b64)) continue;
            if (host.Parts.Any(p => p.RelationshipId == rid)) continue; // idempotent
            byte[] bytes;
            try { bytes = Convert.FromBase64String(b64); }
            catch { continue; }
            var imgPart = host.AddNewPart<ImagePart>(
                string.IsNullOrEmpty(ct) ? "image/png" : ct, rid);
            using var ms = new MemoryStream(bytes);
            imgPart.FeedData(ms);
        }
    }

    // Re-add external hyperlink relationships on a diagram data / drawing part
    // with pinned rIds, so a diagram node's <a:hlinkClick r:id> resolves on
    // replay instead of dangling. Numbered keys {prefix}{k}.rid / .target.
    private static void AttachDiagramHyperlinks(OpenXmlPart host, Dictionary<string, string>? properties, string prefix)
    {
        if (properties == null) return;
        for (int k = 0; ; k++)
        {
            if (!properties.TryGetValue($"{prefix}{k}.rid", out var rid) || string.IsNullOrEmpty(rid))
                break;
            if (!properties.TryGetValue($"{prefix}{k}.target", out var target) || string.IsNullOrEmpty(target))
                continue;
            if (host.HyperlinkRelationships.Any(r => r.Id == rid)) continue; // idempotent
            try { host.AddHyperlinkRelationship(new Uri(target, UriKind.RelativeOrAbsolute), true, rid); }
            catch { /* malformed URI — skip rather than abort the whole add */ }
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc, _filePath);

    public void Save()
    {
        // _doc writes through to _backingStream; force the FileStream buffer
        // out to disk so external readers see the latest bytes immediately.
        if (Modified)
        {
            try { ReconcileSlideMasterIds(); }
            catch { /* best-effort id reconcile */ }
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
    private void AtomicWriteBack()
    {
        if (_packageStream == null || _backingStream == null) return;
        OfficeCli.Core.AtomicPackageWriter.Flush(
            _packageStream, _filePath,
            releaseLock: () => { _backingStream!.Dispose(); _backingStream = null; },
            reopenLock: () => { _backingStream = new FileStream(_filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read); });
    }

    /// <summary>See <see cref="OfficeCli.Handlers.WordHandler.DiscardOnDispose"/> —
    /// atomic-batch rollback: drop the in-memory DOM without serializing.</summary>
    public bool DiscardOnDispose { get; set; }

    public void Dispose()
    {
        if (DiscardOnDispose)
        {
            // Atomic-batch rollback: never serialize the poisoned DOM. Close
            // the backing stream FIRST so the package's dispose-time autosave
            // has nowhere to write. The on-disk file keeps the last flushed
            // (pre-batch) state.
            _backingStream?.Dispose();
            _backingStream = null;
            try { _doc.Dispose(); } catch { /* autosave hit the closed stream — intended */ }
            _packageStream?.Dispose();
            _packageStream = null;
            return;
        }
        if (Modified)
        {
            try { ReconcileSlideMasterIds(); }
            catch { /* best-effort id reconcile */ }
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        try { _doc.Save(); } catch { /* read-only or already disposed */ }
        if (_packageStream != null)
        {
            // Editable: read the in-memory package into the atomic temp BEFORE
            // _doc.Dispose() (the SDK may close the stream it was opened over on
            // dispose), then swap it over the original in one File.Replace.
            try { AtomicWriteBack(); } catch { /* best-effort: original left intact */ }
            _doc.Dispose();
            _packageStream.Dispose();
            _packageStream = null;
            _backingStream?.Dispose();
            _backingStream = null;
        }
        else
        {
            _doc.Dispose();
            _backingStream?.Dispose();
            _backingStream = null;
        }
    }

    // Canonical, RELOAD-STABLE ordering for masters and layouts.
    //
    // The SDK's SlideMasterParts / SlideLayoutParts enumerate the part
    // dictionary, whose order can differ between the moment parts are created
    // (dump replay: prune scaffold + grow from source) and a later reload —
    // sample05 bound `add slide layout=1` to a DIFFERENT part than the one
    // raw-set /slideLayout[1] had just written, chaining the slide to the
    // blank layout instead of the title layout. Order by the presentation's
    // sldMasterIdLst and each master's sldLayoutIdLst (XML document order,
    // stable across save/reload); rel-linked-but-undeclared layouts follow
    // their master's declared ones, ordered by part URI as a stable tiebreak.
    internal static List<SlideMasterPart> MastersInOrder(PresentationPart pp)
    {
        var all = pp.SlideMasterParts.ToList();
        var ordered = new List<SlideMasterPart>();
        var seen = new HashSet<SlideMasterPart>();
        var idList = pp.Presentation?.SlideMasterIdList;
        if (idList != null)
        {
            foreach (var id in idList.Elements<SlideMasterId>())
            {
                var rid = id.RelationshipId?.Value;
                if (string.IsNullOrEmpty(rid)) continue;
                try
                {
                    if (pp.GetPartById(rid) is SlideMasterPart smp && seen.Add(smp))
                        ordered.Add(smp);
                }
                catch { }
            }
        }
        foreach (var m in all.OrderBy(m => m.Uri.OriginalString, StringComparer.Ordinal))
            if (seen.Add(m)) ordered.Add(m);
        return ordered;
    }

    internal static List<SlideLayoutPart> LayoutsInOrder(PresentationPart pp)
    {
        var result = new List<SlideLayoutPart>();
        foreach (var mp in MastersInOrder(pp))
        {
            var seen = new HashSet<SlideLayoutPart>();
            var declared = mp.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>()
                ?? Enumerable.Empty<SlideLayoutId>();
            foreach (var id in declared)
            {
                var rid = id.RelationshipId?.Value;
                if (string.IsNullOrEmpty(rid)) continue;
                try
                {
                    if (mp.GetPartById(rid) is SlideLayoutPart lp && seen.Add(lp))
                        result.Add(lp);
                }
                catch { }
            }
            foreach (var lp in mp.SlideLayoutParts.OrderBy(l => l.Uri.OriginalString, StringComparer.Ordinal))
                if (seen.Add(lp)) result.Add(lp);
        }
        return result;
    }

    // Internal accessors used by PptxBatchEmitter (resource enumeration).
    // Keep the PresentationPart itself private; expose only the counts and
    // a binary getter that the emitter needs.
    internal int SlideMasterCount =>
        _doc.PresentationPart?.SlideMasterParts.Count() ?? 0;
    internal int SlideLayoutCount =>
        _doc.PresentationPart is { } ppLc ? LayoutsInOrder(ppLc).Count : 0;

    /// <summary>
    /// Enumerate ImageParts attached to a slideMaster — one entry per
    /// (rId, content-type, base64 bytes). Used by PptxBatchEmitter to emit
    /// `add-part image` rows so a raw-set'd master XML that references
    /// <c>r:embed="rIdN"</c> on <c>&lt;p:pic&gt;</c> blipFills replays
    /// against an ImagePart with the same rId. Returns an empty list when
    /// the master has no embedded images or the index is out of range.
    /// </summary>
    internal readonly record struct MasterImageInfo(string RelId, string ContentType, string Base64Data);

    internal IReadOnlyList<MasterImageInfo> GetMasterImageParts(int masterIdx)
    {
        var result = new List<MasterImageInfo>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return result;
        var master = masters[masterIdx - 1];
        foreach (var img in master.ImageParts)
        {
            var rid = master.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// A slide master's own ThemePart: its relationship id (as the master XML's
    /// package wires it) and the theme XML. Multi-master decks attach a DISTINCT
    /// theme to each master; the rebuild must re-create each one rather than
    /// sharing the presentation's primary theme (PowerPoint refuses a deck whose
    /// masters share or mis-reference themes). Returns null when the master has no
    /// theme part.
    /// </summary>
    internal (string RelId, string ThemeXml)? GetMasterTheme(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return null;
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return null;
        var master = masters[masterIdx - 1];
        var themePart = master.ThemePart;
        if (themePart?.Theme == null) return null;
        return (master.GetIdOfPart(themePart), themePart.Theme.OuterXml);
    }

    /// <summary>
    /// True when the master's ThemePart is a DIFFERENT part from the
    /// presentation's primary ThemePart. sldMasterIdLst order decides master
    /// enumeration, so master[1] is not necessarily the master that shares the
    /// presentation's theme — a deck whose first-enumerated master carries its
    /// own theme must get an add-part theme on replay or it collapses onto the
    /// scaffold's shared theme (wrong colours for every styleRef on its slides).
    /// </summary>
    internal bool MasterThemeIsDistinct(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return false;
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return false;
        var mt = masters[masterIdx - 1].ThemePart;
        return mt != null && !ReferenceEquals(mt, pp.ThemePart);
    }

    /// <summary>The notes master's own ThemePart (rel id + XML), or null.</summary>
    internal (string RelId, string ThemeXml)? GetNotesMasterTheme()
    {
        var pp = _doc.PresentationPart;
        var nmp = pp?.NotesMasterPart;
        var themePart = nmp?.ThemePart;
        if (nmp == null || themePart?.Theme == null) return null;
        return (nmp.GetIdOfPart(themePart), themePart.Theme.OuterXml);
    }

    /// <summary>
    /// Images attached to a slideMaster's own ThemePart — referenced by an
    /// <c>&lt;a:fmtScheme&gt;&lt;a:fillStyleLst&gt;&lt;a:blipFill&gt;</c> texture
    /// fill via r:embed. The theme XML is re-fed verbatim by the add-part theme
    /// carrier, but its ImageParts live in the theme's own .rels and were never
    /// re-emitted — the rebuilt theme kept a dangling r:embed. Same shape as
    /// <see cref="GetThemeImageParts"/> (which covers the presentation's primary
    /// theme); this covers each master's distinct theme. Empty when none.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetMasterThemeImages(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<MasterImageInfo>();
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return Array.Empty<MasterImageInfo>();
        var themePart = masters[masterIdx - 1].ThemePart;
        return themePart == null ? Array.Empty<MasterImageInfo>() : ReadImagePartInfos(themePart);
    }

    /// <summary>Same as <see cref="GetMasterThemeImages"/> for the notes master's theme.</summary>
    internal IReadOnlyList<MasterImageInfo> GetNotesMasterThemeImages()
    {
        var themePart = _doc.PresentationPart?.NotesMasterPart?.ThemePart;
        return themePart == null ? Array.Empty<MasterImageInfo>() : ReadImagePartInfos(themePart);
    }

    // Shared: enumerate a part's child ImageParts as (rId, content-type, base64).
    private static IReadOnlyList<MasterImageInfo> ReadImagePartInfos(OpenXmlPart host)
    {
        var result = new List<MasterImageInfo>();
        foreach (var idp in host.Parts)
        {
            if (idp.OpenXmlPart is not ImagePart img) continue;
            using var s = img.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(idp.RelationshipId, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// 1-based ordinal of a slide's layout within the same
    /// SlideMasterParts→SlideLayoutParts enumeration that
    /// <see cref="ResolveSlideLayout"/> indexes (its numeric-index match path)
    /// and that the raw-set <c>/slideLayout[N]</c> emission walks. Returns null
    /// when the slide or its layout can't be resolved.
    ///
    /// The batch dump emits this as the slide's `layout=` so replay re-binds to
    /// the EXACT source layout. Emitting the layout NAME is ambiguous: decks
    /// routinely carry several layouts sharing a name (e.g. two "标题幻灯片"
    /// under different masters), and ResolveSlideLayout's name match returns the
    /// first — which can chain the slide to the wrong master and silently drop a
    /// master-level background. The ordinal is unambiguous and stable because
    /// replay reconstructs masters/layouts in this same enumeration order.
    /// </summary>
    internal int? GetSlideLayoutOrdinal(int slideNum)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return null;
        var slideParts = GetSlideParts().ToList();
        if (slideNum < 1 || slideNum > slideParts.Count) return null;
        var layoutPart = slideParts[slideNum - 1].SlideLayoutPart;
        if (layoutPart == null) return null;
        var allLayouts = PowerPointHandler.LayoutsInOrder(pp);
        var idx = allLayouts.IndexOf(layoutPart);
        return idx >= 0 ? idx + 1 : null;
    }

    /// <summary>
    /// Images attached to the presentation's main ThemePart — referenced by an
    /// <c>&lt;a:fmtScheme&gt;&lt;a:fillStyleLst&gt;&lt;a:blipFill&gt;</c> texture
    /// fill via r:embed. The theme XML is raw-set verbatim but its ImageParts are
    /// enumerated separately; without re-emitting them the embed rId dangles and
    /// PowerPoint refuses to open the deck. Same shape as
    /// <see cref="GetMasterImageParts"/>.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetThemeImageParts()
    {
        var result = new List<MasterImageInfo>();
        var theme = _doc.PresentationPart?.ThemePart;
        if (theme == null) return result;
        foreach (var img in theme.ImageParts)
        {
            var rid = theme.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>Same as <see cref="GetMasterImageParts"/> for slideLayouts.</summary>
    internal IReadOnlyList<MasterImageInfo> GetLayoutImageParts(int layoutIdx)
    {
        var result = new List<MasterImageInfo>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var layouts = PowerPointHandler.LayoutsInOrder(pp);
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return result;
        var layout = layouts[layoutIdx - 1];
        foreach (var img in layout.ImageParts)
        {
            var rid = layout.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// Non-image binary parts (ExtendedParts) directly attached to a slideMaster
    /// — chiefly the HD Photo (.wdp) backup layer a master-level decorative
    /// picture references via <c>&lt;a14:imgLayer r:embed&gt;</c>. The master XML
    /// is raw-set verbatim (keeping the source rId), and GetMasterImageParts only
    /// re-creates typed ImageParts, so an hdphoto ExtendedPart was dropped and its
    /// r:embed dangled. Surfaced as companion infos (rId + rel-type + content-type
    /// + ext + bytes) so the emitter pins each via add-part extpart, preserving
    /// the original relationship type. Empty when the master has no such parts.
    /// </summary>
    internal IReadOnlyList<BlipCompanionInfo> GetMasterExtendedParts(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<BlipCompanionInfo>();
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return Array.Empty<BlipCompanionInfo>();
        return ReadExtendedPartInfos(masters[masterIdx - 1]);
    }

    /// <summary>
    /// Custom binary ExtendedParts attached directly to the presentation part —
    /// e.g. Google Slides' ppt/metadata (rel type
    /// http://customschemas.google.com/relationships/presentationmetadata),
    /// referenced by <c>&lt;go:slidesCustomData r:id="rIdN"&gt;</c> inside the
    /// presentation extLst. EmitPresentationExtras replays the extLst verbatim
    /// via raw-set, so without re-pinning the part the r:id dangled and
    /// PowerPoint refused the deck. Surfaced as (rId, relType, contentType, ext,
    /// base64) so the emitter pins each via <c>add-part extpart</c> on
    /// <c>/presentation</c>. Same shape as <see cref="GetMasterExtendedParts"/>.
    /// </summary>
    internal IReadOnlyList<BlipCompanionInfo> GetPresentationExtendedParts()
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<BlipCompanionInfo>();
        return ReadExtendedPartInfos(pp);
    }

    /// <summary>Same as <see cref="GetMasterExtendedParts"/> for slideLayouts.</summary>
    internal IReadOnlyList<BlipCompanionInfo> GetLayoutExtendedParts(int layoutIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<BlipCompanionInfo>();
        var layouts = PowerPointHandler.LayoutsInOrder(pp);
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return Array.Empty<BlipCompanionInfo>();
        return ReadExtendedPartInfos(layouts[layoutIdx - 1]);
    }

    /// <summary>
    /// External (TargetMode="External") IMAGE relationships on a slideMaster —
    /// a master picture can LINK to an external image (<a:blip r:link="rIdN">,
    /// TargetMode=External) rather than embed it. The master XML is raw-set
    /// verbatim (keeping r:link="rIdN"), but GetMasterImageParts only re-creates
    /// embedded ImageParts, so the external relationship was dropped and the
    /// rebuilt master's r:link dangled. Surfaced as (rId, relationship-type, uri)
    /// so the emitter pins each via add-part extrel. (The hyperlink carrier
    /// already covers .../hyperlink external rels; this covers the .../image
    /// external links.) Empty when the master links no external images.
    /// </summary>
    internal IReadOnlyList<(string RelId, string RelType, string Uri)> GetMasterExternalImageLinks(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<(string, string, string)>();
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return Array.Empty<(string, string, string)>();
        return ReadExternalImageLinks(masters[masterIdx - 1]);
    }

    /// <summary>Same as <see cref="GetMasterExternalImageLinks"/> for slideLayouts.</summary>
    internal IReadOnlyList<(string RelId, string RelType, string Uri)> GetLayoutExternalImageLinks(int layoutIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<(string, string, string)>();
        var layouts = PowerPointHandler.LayoutsInOrder(pp);
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return Array.Empty<(string, string, string)>();
        return ReadExternalImageLinks(layouts[layoutIdx - 1]);
    }

    private static IReadOnlyList<(string RelId, string RelType, string Uri)> ReadExternalImageLinks(OpenXmlPart host)
    {
        var result = new List<(string, string, string)>();
        foreach (var rel in host.ExternalRelationships)
        {
            // Image external links (r:link). Skip hyperlinks (carried separately).
            if (rel.RelationshipType.EndsWith("/image", StringComparison.Ordinal))
                result.Add((rel.Id, rel.RelationshipType, rel.Uri.OriginalString));
        }
        return result;
    }

    private static IReadOnlyList<BlipCompanionInfo> ReadExtendedPartInfos(OpenXmlPart host)
    {
        var result = new List<BlipCompanionInfo>();
        foreach (var idp in host.Parts)
        {
            // Arbitrary binary blobs reached by a custom or non-typed-image
            // relationship: ExtendedPart (hdphoto / Google metadata / …) plus
            // embedded OLE objects (EmbeddedObjectPart, rel .../oleObject) and
            // embedded packages (EmbeddedPackagePart, embedded .docx/.xlsx).
            // A master/layout can host an <p:oleObj r:id> (e.g. clip-art) whose
            // part is one of the latter two; the raw-set master XML keeps the
            // r:id, so the part must be re-pinned or the reference dangles and
            // PowerPoint refuses the deck. ImageParts are carried separately
            // (GetMasterImageParts), so they are intentionally excluded here.
            OpenXmlPart? blob = idp.OpenXmlPart switch
            {
                ExtendedPart e => e,
                EmbeddedObjectPart o => o,
                EmbeddedPackagePart p => p,
                _ => null
            };
            if (blob == null) continue;
            using var s = blob.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            var ext = System.IO.Path.GetExtension(blob.Uri.OriginalString);
            if (string.IsNullOrEmpty(ext)) ext = ".bin";
            result.Add(new BlipCompanionInfo(
                idp.RelationshipId, blob.RelationshipType, blob.ContentType, ext,
                Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// External (TargetMode="External") hyperlink relationships on a slideLayout.
    /// The layout XML is replayed via raw-set carrying <c>&lt;a:hlinkClick r:id="rIdN"/&gt;</c>,
    /// but the referenced relationship is external (a URL, not an embedded part),
    /// so the ImagePart carrier never re-creates it — the renumbered rebuilt
    /// layout's .rels lost rIdN and PowerPoint refused the file. Surfaced as
    /// (rId, target) pairs so PptxBatchEmitter can emit an `add-part hyperlink`
    /// row that pins each id before the layout raw-set replace.
    /// </summary>
    internal IReadOnlyList<(string RelId, string Target)> GetLayoutExternalHyperlinks(int layoutIdx)
    {
        var result = new List<(string, string)>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var layouts = PowerPointHandler.LayoutsInOrder(pp);
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return result;
        var layout = layouts[layoutIdx - 1];
        foreach (var rel in layout.HyperlinkRelationships)
        {
            if (rel.IsExternal)
                result.Add((rel.Id, rel.Uri.OriginalString));
        }
        return result;
    }

    /// <summary>Same as <see cref="GetLayoutExternalHyperlinks"/> for a slideMaster.
    /// A master shape (e.g. a footer / "designed by" credit) can carry
    /// <c>&lt;a:hlinkClick r:id="rIdN"/&gt;</c> to an external URL. The master is
    /// replayed via raw-set which keeps the reference, but the ImagePart carrier
    /// never re-creates an external hyperlink rel, so the rebuilt master's .rels
    /// lost rIdN and PowerPoint refused the whole deck (0x80070570 OPC corrupt).
    /// Surfaced as (rId, target) so the emitter pins each id via an
    /// <c>add-part hyperlink</c> row before the master raw-set replace.</summary>
    internal IReadOnlyList<(string RelId, string Target)> GetMasterExternalHyperlinks(int masterIdx)
    {
        var result = new List<(string, string)>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return result;
        foreach (var rel in masters[masterIdx - 1].HyperlinkRelationships)
        {
            if (rel.IsExternal)
                result.Add((rel.Id, rel.Uri.OriginalString));
        }
        return result;
    }

    /// <summary>
    /// UserDefinedTags parts (programmability metadata, <c>&lt;p:tagLst&gt;</c>)
    /// attached to a slideLayout, surfaced as (rId, verbatim tag XML) pairs.
    /// The layout XML is replayed via raw-set carrying
    /// <c>&lt;p:custDataLst&gt;&lt;p:tags r:id="rIdN"/&gt;</c>, but the tags part
    /// lives in the layout's own .rels (enumerated separately) and was never
    /// re-emitted — the rebuilt layout's <c>r:id="rIdN"</c> then dangled and
    /// PowerPoint refused the whole deck (0x80070570 OPC corrupt). Emitting an
    /// <c>add-part tags</c> row that pins each source rId before the layout
    /// raw-set replace makes the reference resolve. Same shape as
    /// <see cref="GetLayoutImageParts"/>.
    /// </summary>
    internal IReadOnlyList<(string RelId, string TagXml)> GetLayoutTagParts(int layoutIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<(string, string)>();
        var layouts = PowerPointHandler.LayoutsInOrder(pp);
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return Array.Empty<(string, string)>();
        return ReadTagParts(layouts[layoutIdx - 1]);
    }

    /// <summary>Same as <see cref="GetLayoutTagParts"/> for a slideMaster.</summary>
    internal IReadOnlyList<(string RelId, string TagXml)> GetMasterTagParts(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<(string, string)>();
        var masters = MastersInOrder(pp);
        if (masterIdx < 1 || masterIdx > masters.Count) return Array.Empty<(string, string)>();
        return ReadTagParts(masters[masterIdx - 1]);
    }

    private static IReadOnlyList<(string RelId, string TagXml)> ReadTagParts(OpenXmlPartContainer host)
    {
        var result = new List<(string, string)>();
        foreach (var idp in host.Parts)
        {
            if (idp.OpenXmlPart is not UserDefinedTagsPart tagPart) continue;
            using var s = tagPart.GetStream(FileMode.Open, FileAccess.Read);
            using var sr = new StreamReader(s);
            result.Add((idp.RelationshipId, sr.ReadToEnd()));
        }
        return result;
    }

    /// <summary>
    /// Same as <see cref="GetMasterImageParts"/> for a slide's NotesSlidePart.
    /// Used by PptxBatchEmitter.EmitNotes so a notesSlide raw-set replace that
    /// references <c>r:embed="rIdN"</c> on a <c>&lt;p:pic&gt;</c> blipFill
    /// (image pasted into speaker notes) replays against an ImagePart with
    /// the same rId. Without this the post-replay notesSlide carries a
    /// dangling rId and PowerPoint shows a broken picture placeholder.
    /// Returns an empty list when the slide has no notesSlide or no embedded
    /// images.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetNoteSlideImageParts(int slideIdx)
    {
        var result = new List<MasterImageInfo>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var notesPart = slideParts[slideIdx - 1].NotesSlidePart;
        if (notesPart == null) return result;
        foreach (var img in notesPart.ImageParts)
        {
            var rid = notesPart.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// External (TargetMode="External") hyperlink relationships on a slide's
    /// NotesSlidePart — same shape as <see cref="GetLayoutExternalHyperlinks"/>.
    /// The notesSlide XML is replayed via raw-set carrying
    /// <c>&lt;a:hlinkClick r:id="rIdN"/&gt;</c> (a URL in the speaker notes), but
    /// the external relationship is not an embedded part, so the ImagePart carrier
    /// never re-creates it — the rebuilt notesSlide's <c>r:id="rIdN"</c> dangled
    /// and PowerPoint refused the whole deck (OPC corrupt). Surfaced as
    /// (rId, target) pairs so EmitNotes can pin each id via an
    /// <c>add-part hyperlink</c> row before the notes raw-set replace.
    /// </summary>
    internal IReadOnlyList<(string RelId, string Target)> GetNoteSlideExternalHyperlinks(int slideIdx)
    {
        var result = new List<(string, string)>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var notesPart = slideParts[slideIdx - 1].NotesSlidePart;
        if (notesPart == null) return result;
        foreach (var rel in notesPart.HyperlinkRelationships)
        {
            if (rel.IsExternal)
                result.Add((rel.Id, rel.Uri.OriginalString));
        }
        return result;
    }

    /// <summary>
    /// External (TargetMode="External") hyperlink relationships on a slide whose
    /// relationship id is one of <paramref name="relIds"/>. A table cell (and
    /// other raw-passthrough content) is replayed via verbatim txBodyRaw that
    /// keeps <c>&lt;a:hlinkClick r:id="rIdN"&gt;</c> pointing at a URL; the typed
    /// emit never re-creates that external relationship, so the rebuilt slide
    /// kept a dangling rId. Surfaced as (rId, target) so the emitter pins each
    /// via add-part hyperlink. Scoped to the requested rIds (the ones the raw
    /// body actually references) to avoid re-creating links the typed `link=`
    /// path already rebuilt. Mirrors <see cref="GetNoteSlideExternalHyperlinks"/>.
    /// </summary>
    internal IReadOnlyList<(string RelId, string Target)> GetSlideExternalHyperlinksByRelId(
        int slideIdx, IReadOnlyCollection<string> relIds)
    {
        var result = new List<(string, string)>();
        if (relIds.Count == 0) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var slide = slideParts[slideIdx - 1];
        var wanted = new HashSet<string>(relIds, StringComparer.Ordinal);
        foreach (var rel in slide.HyperlinkRelationships)
        {
            if (rel.IsExternal && wanted.Contains(rel.Id))
                result.Add((rel.Id, rel.Uri.OriginalString));
        }
        return result;
    }

    /// <summary>
    /// Internal slide-jump relationships on a slide whose id is one of
    /// <paramref name="relIds"/>: a run's <c>&lt;a:hlinkClick r:id="rIdN"
    /// action="ppaction://hlinksldjump"&gt;</c> targets ANOTHER slide via a
    /// relationship of type .../slide. When such a link lives in a table cell's
    /// verbatim txBodyRaw, the typed slide-jump path (DeferSlideJumpLink →
    /// link=slide[N]) never fires, so the relationship is not re-created and the
    /// rebuilt slide's r:id="rIdN" dangles — PowerPoint then refuses the deck
    /// (0x80070570). Surfaced as (rId, targetSlideOrdinal) — the 1-based ordinal
    /// of the target slide within <see cref="GetSlideParts"/> — so the emitter can
    /// pin the rId to the rebuilt target slide AFTER every slide exists.
    /// </summary>
    internal IReadOnlyList<(string RelId, int TargetOrdinal)> GetSlideInternalSlideJumpRels(
        int slideIdx, IReadOnlyCollection<string> relIds)
    {
        var result = new List<(string, int)>();
        if (relIds.Count == 0) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var slide = slideParts[slideIdx - 1];
        var wanted = new HashSet<string>(relIds, StringComparer.Ordinal);
        foreach (var idp in slide.Parts)
        {
            if (!wanted.Contains(idp.RelationshipId)) continue;
            if (idp.OpenXmlPart is SlidePart tgt)
            {
                var ord = slideParts.IndexOf(tgt);
                if (ord >= 0) result.Add((idp.RelationshipId, ord + 1));
            }
        }
        return result;
    }

    /// <summary>
    /// ImageParts on a slide whose relationship id is one of <paramref name="relIds"/>.
    /// Used by PptxBatchEmitter.EmitRawSlideBgSlice so a slide-level
    /// <c>&lt;p:bg&gt;&lt;p:bgPr&gt;&lt;a:blipFill&gt;&lt;a:blip r:embed="rIdN"&gt;</c>
    /// (background image) raw-set replays against an ImagePart carrying the
    /// SAME source rId. The bg slice is emitted verbatim via raw-set, so its
    /// <c>r:embed="rIdN"</c> only resolves if a matching ImagePart with that
    /// pinned rId exists on the rebuilt slide. We scope to the bg-referenced
    /// rIds only — slide pictures already round-trip through the typed
    /// <c>add picture</c> (fresh rId) path, so re-creating every ImagePart
    /// here would double-create them. Returns an empty list when the slide
    /// is out of range or carries none of the requested rIds.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetSlideImagePartsByRelId(
        int slideIdx, IReadOnlyCollection<string> relIds)
    {
        var result = new List<MasterImageInfo>();
        if (relIds.Count == 0) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var slide = slideParts[slideIdx - 1];
        var wanted = new HashSet<string>(relIds, StringComparer.Ordinal);
        foreach (var img in slide.ImageParts)
        {
            var rid = slide.GetIdOfPart(img);
            if (!wanted.Contains(rid)) continue;
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// Images a slide references from RAW-passthrough text/bullet contexts that
    /// the typed emit never re-creates: picture-bullet glyphs
    /// (<c>&lt;a:buBlip&gt;&lt;a:blip&gt;</c>) and image text-fills
    /// (<c>&lt;a:defRPr&gt;/&lt;a:rPr&gt;/&lt;a:lvlNpPr&gt;…&lt;a:blipFill&gt;&lt;a:blip&gt;</c>,
    /// where the glyph outlines are filled with an image). Both round-trip
    /// verbatim via bulletRaw / lstStyleRaw / defRPrRaw keeping <c>r:embed="rIdN"</c>,
    /// but the slide ImagePart was never re-emitted (the typed `add picture` path
    /// only covers <p:pic>, and a shape's own <p:spPr> blipFill is handled by the
    /// image=true carrier), so the rebuilt slide dangled. Surfaced as
    /// (rId, content-type, base64) with the SOURCE rId pinned so the emitter
    /// re-creates each via an add-part image row BEFORE shapes are added
    /// (claiming the source rId before AddPicture auto-assigns around it).
    /// Excludes <p:pic> main blips and shape <p:spPr> fills — those round-trip
    /// elsewhere — so it never double-creates a typed-emitted image.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetSlideBulletImageParts(int slideIdx)
    {
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return Array.Empty<MasterImageInfo>();
        var spTree = GetSlide(slideParts[slideIdx - 1]).CommonSlideData?.ShapeTree;
        if (spTree == null) return Array.Empty<MasterImageInfo>();
        var rids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var blip in spTree.Descendants<DocumentFormat.OpenXml.Drawing.Blip>())
        {
            var embed = blip.Embed?.Value;
            if (string.IsNullOrEmpty(embed)) continue;
            var parent = blip.Parent;
            if (parent == null) continue;
            if (parent.LocalName == "buBlip")
            {
                rids.Add(embed); // picture-bullet glyph
            }
            else if (parent.LocalName == "blipFill")
            {
                var gp = parent.Parent?.LocalName;
                // spPr → shape image-fill (image=true carrier handles it);
                // pic → typed picture (add picture handles it). Anything else
                // (defRPr / rPr / lvlNpPr / lstStyle text-fill) is raw-only.
                if (gp != "spPr" && gp != "pic") rids.Add(embed);
            }
        }
        return rids.Count == 0 ? Array.Empty<MasterImageInfo>() : GetSlideImagePartsByRelId(slideIdx, rids);
    }

    /// <summary>
    /// Enumerate <c>r:embed</c> / <c>r:link</c> attribute values referenced by
    /// the source notesSlide XML. Used by PptxBatchEmitter.EmitNotes to detect
    /// rIds that the typed Add/Set surface cannot reproduce (anything not an
    /// ImagePart enumerated by <see cref="GetNoteSlideImageParts"/>). When such
    /// orphan rIds exist, the emitter surfaces a
    /// <c>notes_unresolved_rid</c> warning so callers know the post-replay
    /// notesSlide may have dangling references that PowerPoint will render
    /// as broken placeholders.
    /// </summary>
    internal IReadOnlyList<string> GetNoteSlideExternalRelIds(int slideIdx)
    {
        var result = new List<string>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var notesPart = slideParts[slideIdx - 1].NotesSlidePart;
        if (notesPart == null) return result;
        var xml = notesPart.NotesSlide?.OuterXml;
        if (string.IsNullOrEmpty(xml)) return result;
        var rx = new Regex(@"r:(?:embed|link|id)=""([^""]+)""");
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match m in rx.Matches(xml))
        {
            var rid = m.Groups[1].Value;
            if (seen.Add(rid)) result.Add(rid);
        }
        return result;
    }

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

    /// <summary>
    /// Per-slide SmartArt info for PptxBatchEmitter passthrough. Returns
    /// one entry per <p:graphicFrame> on the slide that carries a
    /// <dgm:relIds> child (= SmartArt host frame). Each entry includes the
    /// source's four rIds and the four diagram parts' XML so the emitter
    /// can issue an `add-part smartart` + four `raw-set` rows that
    /// round-trip byte-equal.
    /// </summary>
    internal readonly record struct SmartArtInfo(
        string GraphicFrameXml,
        string DataRelId,
        string LayoutRelId,
        string ColorsRelId,
        string QuickStyleRelId,
        string DataXml,
        string LayoutXml,
        string ColorsXml,
        string QuickStyleXml,
        string? DrawingXml,
        string? DrawingRelId,
        // Images referenced by the data part and the DSP drawing part via their
        // OWN .rels (picture-in-diagram blipFills). Both parts are recreated
        // empty by add-part smartart, so without re-attaching these ImageParts
        // with pinned rIds their r:embed references dangle and PowerPoint refuses
        // the deck (0x80070570). Empty when the SmartArt carries no pictures.
        IReadOnlyList<MasterImageInfo> DataImages,
        IReadOnlyList<MasterImageInfo> DrawingImages,
        // External hyperlink relationships on the data part and the DSP drawing
        // part's OWN .rels — a diagram node can carry an <a:hlinkClick r:id> to
        // an external URL. add-part smartart recreates both parts empty, so
        // without re-adding these external relationships with pinned rIds their
        // r:id dangles and PowerPoint refuses the deck (0x80070570). Empty when
        // the SmartArt carries no hyperlinks.
        IReadOnlyList<(string RelId, string Target)> DataHyperlinks,
        IReadOnlyList<(string RelId, string Target)> DrawingHyperlinks,
        // spTree-rooted xpath of the graphicFrame's PARENT container. Top-level
        // SmartArt -> "/p:sld/p:cSld/p:spTree"; a group-nested SmartArt ->
        // ".../p:grpSp[K]..." so the emitter re-appends it INSIDE the group,
        // preserving the group's transform/scaling. Without this a group-nested
        // SmartArt was relocated to the slide top level and rendered at the
        // wrong (unscaled) size — overflowing the slide.
        string ParentXpath,
        // cNvPr id of the first following sibling in the SOURCE parent that is a
        // "stable" main-walk element (plain shape / connector / group /
        // chart|table graphicFrame — one guaranteed present in the rebuilt slide
        // by the time the SmartArt pass runs). The emitter inserts the SmartArt's
        // graphicFrame BEFORE that element so z-order is preserved. null when the
        // SmartArt is the last child (or is only followed by deferred/exotic
        // elements) → the emitter appends, as before. Without this a SmartArt
        // that sat BEHIND a later shape was appended last and rendered ON TOP,
        // occluding that shape (bnc880763).
        string? InsertBeforeShapeId,
        // Positional sibling xpath segment (e.g. "/p:sp[1]") for group-nested
        // SmartArt — group-descendant cNvPr ids are reassigned on replay, so
        // the @id anchor form cannot match there.
        string? InsertBeforeSegment = null);

    // External (TargetMode="External") hyperlink relationships on a part's own
    // .rels, surfaced as (rId, target-uri). Mirrors GetLayoutExternalHyperlinks
    // for any OpenXmlPartContainer host (diagram data / drawing parts).
    private static IReadOnlyList<(string RelId, string Target)> ReadExternalHyperlinksOf(OpenXmlPart host)
    {
        var result = new List<(string, string)>();
        foreach (var rel in host.HyperlinkRelationships)
            if (rel.IsExternal) result.Add((rel.Id, rel.Uri.OriginalString));
        return result;
    }

    // Enumerate the ImageParts directly attached to a part (its own .rels),
    // surfaced as (rId, content-type, base64). Mirrors GetMasterImageParts but
    // for any OpenXmlPartContainer host (diagram data / drawing parts).
    private static IReadOnlyList<MasterImageInfo> ReadImagePartsOf(OpenXmlPart host)
    {
        var result = new List<MasterImageInfo>();
        foreach (var idp in host.Parts)
        {
            if (idp.OpenXmlPart is not ImagePart img) continue;
            using var s = img.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(idp.RelationshipId, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    // A sound relationship referenced from a slide's <p:timing> tree
    // (<p:sndTgt r:embed>) or its <p:transition> (<p:snd r:embed>). The timing /
    // transition XML is round-tripped verbatim via raw-set, so its literal rId
    // must resolve on replay — but those rels are NOT part of the <p:pic> media
    // pass, so without re-creating them the r:embed dangles and PowerPoint
    // refuses the deck (0x80070570). Carry the bytes + content type + pinned rId.
    internal readonly record struct TimingAudioRel(
        string RelId, string ContentType, string Extension, byte[] Data);

    internal IReadOnlyList<TimingAudioRel> GetTimingAudioRels(int slideIdx)
    {
        var result = new List<TimingAudioRel>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];

        string slideXml;
        try
        {
            using var s = slidePart.GetStream(FileMode.Open, FileAccess.Read);
            using var r = new StreamReader(s);
            slideXml = r.ReadToEnd();
        }
        catch { return result; }

        // Collect rIds referenced by sound elements (timing sndTgt + transition snd).
        var rids = new HashSet<string>(StringComparer.Ordinal);
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
            slideXml, @"<p:snd(?:Tgt)?\b[^>]*\br:embed=""([^""]+)"""))
        {
            rids.Add(m.Groups[1].Value);
        }
        if (rids.Count == 0) return result;

        // Sound rels are DATA-PART reference relationships (media), not regular
        // part relationships — GetPartById throws on them. Resolve the DataPart
        // via the slide's DataPartReferenceRelationships keyed by rId.
        var byId = new Dictionary<string, DocumentFormat.OpenXml.Packaging.DataPart>(StringComparer.Ordinal);
        foreach (var dpr in slidePart.DataPartReferenceRelationships)
            if (dpr.Id is { } id && !byId.ContainsKey(id)) byId[id] = dpr.DataPart;

        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rid in rids)
        {
            if (!seen.Add(rid)) continue;
            if (!byId.TryGetValue(rid, out var dataPart) || dataPart == null) continue;
            try
            {
                byte[] bytes;
                using (var ps = dataPart.GetStream(FileMode.Open, FileAccess.Read))
                using (var ms = new MemoryStream())
                {
                    ps.CopyTo(ms);
                    bytes = ms.ToArray();
                }
                if (bytes.Length == 0) continue;
                var ct = string.IsNullOrEmpty(dataPart.ContentType) ? "audio/wav" : dataPart.ContentType;
                var ext = Path.GetExtension(dataPart.Uri?.OriginalString ?? "").ToLowerInvariant();
                if (string.IsNullOrEmpty(ext)) ext = ".wav";
                result.Add(new TimingAudioRel(rid, ct, ext, bytes));
            }
            catch { /* dangling in source too — skip */ }
        }
        return result;
    }

    // ActiveX controls on a slide. Producers put a <p:controls>
    // block in <p:cSld> (sibling of the shape tree) holding one
    // <mc:AlternateContent> per control (Choice = <p:control> VML fallback,
    // Fallback = <p:control><p:pic> the rendered image). The control references
    // an activeX{N}.xml part (rel type .../control), which in turn references a
    // .bin blob via its OWN .rels. The fallback pic references a WMF image.
    // NodeBuilder does not surface any of this, so the whole block + its parts
    // were dropped on dump → the slide replayed blank. Carry the controls XML
    // verbatim plus every referenced part (activeX xml + its bin child + WMF
    // images) with pinned rIds.
    internal readonly record struct ActiveXControlPart(string ControlRid, byte[] Xml, string? BinRid, byte[]? Bin);
    internal readonly record struct ActiveXImage(string Rid, string ContentType, string Ext, byte[] Bytes);

    internal (string? ControlsXml,
              IReadOnlyList<ActiveXControlPart> Controls,
              IReadOnlyList<ActiveXImage> Images) GetActiveXOnSlide(int slideIdx)
    {
        var noControls = ((string?)null,
            (IReadOnlyList<ActiveXControlPart>)Array.Empty<ActiveXControlPart>(),
            (IReadOnlyList<ActiveXImage>)Array.Empty<ActiveXImage>());
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return noControls;
        var slidePart = parts[slideIdx - 1];

        string slideXml;
        try
        {
            using var s = slidePart.GetStream(FileMode.Open, FileAccess.Read);
            using var r = new StreamReader(s);
            slideXml = r.ReadToEnd();
        }
        catch { return noControls; }

        var cm = System.Text.RegularExpressions.Regex.Match(slideXml,
            @"<p:controls>.*?</p:controls>", System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!cm.Success) return noControls;
        var controlsXml = cm.Value;

        const string ControlRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
        const string ImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        // rIds referenced inside <p:controls> (control r:id + pic blip r:embed).
        var refIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(
            controlsXml, @"r:(?:id|embed)=""(rId\d+)"""))
            refIds.Add(m.Groups[1].Value);

        var controls = new List<ActiveXControlPart>();
        var images = new List<ActiveXImage>();
        var seenImg = new HashSet<string>(StringComparer.Ordinal);

        foreach (var rel in slidePart.Parts)
        {
            if (!refIds.Contains(rel.RelationshipId)) continue;
            try
            {
                if (rel.OpenXmlPart.RelationshipType == ControlRelType)
                {
                    byte[] axXml;
                    using (var ps = rel.OpenXmlPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (var ms = new MemoryStream()) { ps.CopyTo(ms); axXml = ms.ToArray(); }
                    // The control's .bin child (activeXControlBinary), if any.
                    string? binRid = null; byte[]? bin = null;
                    var binPair = rel.OpenXmlPart.Parts.FirstOrDefault();
                    if (binPair.OpenXmlPart != null)
                    {
                        binRid = binPair.RelationshipId;
                        using var bs = binPair.OpenXmlPart.GetStream(FileMode.Open, FileAccess.Read);
                        using var bms = new MemoryStream(); bs.CopyTo(bms); bin = bms.ToArray();
                    }
                    controls.Add(new ActiveXControlPart(rel.RelationshipId, axXml, binRid, bin));
                }
                else if (rel.OpenXmlPart.RelationshipType == ImageRelType)
                {
                    if (!seenImg.Add(rel.RelationshipId)) continue;
                    byte[] bytes;
                    using (var ps = rel.OpenXmlPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (var ms = new MemoryStream()) { ps.CopyTo(ms); bytes = ms.ToArray(); }
                    var ct = string.IsNullOrEmpty(rel.OpenXmlPart.ContentType) ? "image/x-wmf" : rel.OpenXmlPart.ContentType;
                    var ext = Path.GetExtension(rel.OpenXmlPart.Uri?.OriginalString ?? "");
                    if (string.IsNullOrEmpty(ext)) ext = ".wmf";
                    images.Add(new ActiveXImage(rel.RelationshipId, ct, ext, bytes));
                }
            }
            catch { /* skip malformed */ }
        }

        if (controls.Count == 0) return noControls;
        return (controlsXml, controls, images);
    }

    internal IReadOnlyList<SmartArtInfo> GetSmartArtsOnSlide(int slideIdx)
    {
        var result = new List<SmartArtInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];
        var slide = GetSlide(slidePart);
        var spTree = slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return result;

        var ns = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        foreach (var gf in spTree.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>())
        {
            var relIds = gf.Descendants().FirstOrDefault(e =>
                e.LocalName == "relIds" && e.NamespaceUri == ns);
            if (relIds == null) continue;

            string? dRid = null, lRid = null, cRid = null, qRid = null;
            foreach (var a in relIds.GetAttributes())
            {
                var ln = a.LocalName;
                var v = a.Value;
                if (ln == "dm") dRid = v;
                else if (ln == "lo") lRid = v;
                else if (ln == "cs") cRid = v;
                else if (ln == "qs") qRid = v;
            }
            if (dRid == null || lRid == null || cRid == null || qRid == null) continue;

            string? xmlFor(string rid)
            {
                try
                {
                    var part = slidePart.GetPartById(rid);
                    if (part is DiagramDataPart d) return d.DataModelRoot?.OuterXml;
                    if (part is DiagramLayoutDefinitionPart l) return l.LayoutDefinition?.OuterXml;
                    if (part is DiagramColorsPart c) return c.ColorsDefinition?.OuterXml;
                    if (part is DiagramStylePart s) return s.StyleDefinition?.OuterXml;
                }
                catch { }
                return null;
            }

            var dXml = xmlFor(dRid);
            var lXml = xmlFor(lRid);
            var cXml = xmlFor(cRid);
            var qXml = xmlFor(qRid);
            if (dXml == null || lXml == null || cXml == null || qXml == null) continue;

            // The data part references a 5th part — the DSP cached-drawing
            // part — via <dsp:dataModelExt relId="..."> in the data XML. That
            // relId resolves against the SLIDE part's relationships (the
            // drawing part is a slide-level part of type
            // .../2007/relationships/diagramDrawing, sibling to the
            // data/layout/colors/qs rels — NOT a child of the data part).
            // PowerPoint refuses the file if that relId dangles, so carry the
            // drawing XML + the relId. Leave both null when absent (older /
            // simpler SmartArt without a cached drawing) → keep behavior.
            string? drawingXml = null, drawingRelId = null;
            IReadOnlyList<MasterImageInfo> dataImages = Array.Empty<MasterImageInfo>();
            IReadOnlyList<MasterImageInfo> drawingImages = Array.Empty<MasterImageInfo>();
            IReadOnlyList<(string, string)> dataHlinks = Array.Empty<(string, string)>();
            IReadOnlyList<(string, string)> drawingHlinks = Array.Empty<(string, string)>();
            try
            {
                if (slidePart.GetPartById(dRid) is DiagramDataPart ddp)
                {
                    // Pictures embedded in the diagram (point-level blipFills)
                    // live as ImageParts on the data part's own .rels.
                    try { dataImages = ReadImagePartsOf(ddp); } catch { }
                    // External hyperlinks on diagram nodes live as hyperlink rels.
                    try { dataHlinks = ReadExternalHyperlinksOf(ddp); } catch { }
                    const string dspNs = "http://schemas.microsoft.com/office/drawing/2008/diagram";
                    var ext = ddp.DataModelRoot?.Descendants().FirstOrDefault(e =>
                        e.LocalName == "dataModelExt" && e.NamespaceUri == dspNs);
                    if (ext != null)
                    {
                        foreach (var a in ext.GetAttributes())
                            if (a.LocalName == "relId") { drawingRelId = a.Value; break; }
                    }
                    if (!string.IsNullOrEmpty(drawingRelId))
                    {
                        try
                        {
                            if (slidePart.GetPartById(drawingRelId) is DiagramPersistLayoutPart drawingPart)
                            {
                                // The DSP cached drawing re-references the same
                                // pictures for rendering via its own .rels.
                                try { drawingImages = ReadImagePartsOf(drawingPart); } catch { }
                                try { drawingHlinks = ReadExternalHyperlinksOf(drawingPart); } catch { }
                                using var s = drawingPart.GetStream(FileMode.Open, FileAccess.Read);
                                using var r = new StreamReader(s);
                                drawingXml = r.ReadToEnd();
                                // Strip XML prolog so emit/replay re-adds it
                                // uniformly (WriteDiagramPartXml re-prepends).
                                int lt = drawingXml.IndexOf('<', StringComparison.Ordinal);
                                int decl = drawingXml.IndexOf("<?xml", StringComparison.Ordinal);
                                if (decl == 0)
                                {
                                    int end = drawingXml.IndexOf("?>", StringComparison.Ordinal);
                                    if (end >= 0) drawingXml = drawingXml.Substring(end + 2).TrimStart();
                                }
                                else if (lt > 0) drawingXml = drawingXml.Substring(lt);
                            }
                        }
                        catch { drawingXml = null; }
                    }
                }
            }
            catch { drawingXml = null; drawingRelId = null; }
            if (drawingXml == null || drawingRelId == null) { drawingXml = null; drawingRelId = null; }

            // Parent-container xpath: walk group ancestors so a group-nested
            // SmartArt round-trips inside its group (keeps the group scaling)
            // rather than being relocated to the slide top level.
            var grpChain = new List<string>();
            for (var cur = gf.Parent; cur != null && cur != spTree; cur = cur.Parent)
            {
                if (cur is DocumentFormat.OpenXml.Presentation.GroupShape gs)
                {
                    int gIdx = 1;
                    foreach (var sib in gs.Parent!.Elements<DocumentFormat.OpenXml.Presentation.GroupShape>())
                    {
                        if (ReferenceEquals(sib, gs)) break;
                        gIdx++;
                    }
                    grpChain.Insert(0, $"/p:grpSp[{gIdx}]");
                }
            }
            var parentXpath = "/p:sld/p:cSld/p:spTree" + string.Concat(grpChain);

            // Z-order anchor: the first following sibling in the SAME parent that
            // is a "stable" element already present when the SmartArt pass runs.
            // A SmartArt is emitted last (add-part + raw-set append), so without
            // this it always lands on top; anchoring an insert before the next
            // stable sibling preserves the source z-order.
            string? insertBeforeShapeId = null;
            string? insertBeforeSegment = null;
            foreach (var sib in gf.ElementsAfter())
            {
                if (!IsStableZOrderAnchor(sib)) continue;
                if (grpChain.Count > 0)
                {
                    // Group-nested: replay REASSIGNS group-descendant cNvPr ids
                    // (CONSISTENCY(group-id-autoassign)), so an @id anchor
                    // matches nothing and the SmartArt silently never lands
                    // (smartart-groupshape). Use the sibling's positional form
                    // instead — same-tag ordinal within the group, stable
                    // because the group's children replay in source order.
                    int ord = 1;
                    foreach (var prev in sib.Parent!.ChildElements)
                    {
                        if (ReferenceEquals(prev, sib)) break;
                        if (prev.LocalName == sib.LocalName) ord++;
                    }
                    insertBeforeSegment = $"/p:{sib.LocalName}[{ord}]";
                    break;
                }
                var cnv = sib.Descendants().FirstOrDefault(e =>
                    e.LocalName == "cNvPr"
                    && e.NamespaceUri == "http://schemas.openxmlformats.org/presentationml/2006/main");
                var idAttr = cnv?.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
                if (idAttr?.Value is { Length: > 0 } idVal) { insertBeforeShapeId = idVal; break; }
            }

            result.Add(new SmartArtInfo(
                GraphicFrameXml: gf.OuterXml,
                DataRelId: dRid, LayoutRelId: lRid, ColorsRelId: cRid, QuickStyleRelId: qRid,
                DataXml: dXml, LayoutXml: lXml, ColorsXml: cXml, QuickStyleXml: qXml,
                DrawingXml: drawingXml, DrawingRelId: drawingRelId,
                DataImages: dataImages, DrawingImages: drawingImages,
                DataHyperlinks: dataHlinks, DrawingHyperlinks: drawingHlinks,
                ParentXpath: parentXpath,
                InsertBeforeShapeId: insertBeforeShapeId,
                InsertBeforeSegment: insertBeforeSegment));
        }
        return result;
    }

    /// <summary>
    /// Is <paramref name="el"/> a shape-tree element guaranteed to already exist
    /// in the rebuilt slide when the SmartArt emit pass runs (so a SmartArt can
    /// safely be inserted BEFORE it to preserve z-order)? True for plain shapes,
    /// connectors, groups, and chart/table graphicFrames — all emitted in the
    /// main child walk (or the connector flush) before the SmartArt pass. False
    /// for elements that are themselves emitted late or exotically: media
    /// (&lt;p:pic&gt; with video/audio), OLE (graphicFrame with &lt;p:oleObj&gt;),
    /// another SmartArt (graphicFrame with dgm:relIds), and mc:AlternateContent
    /// (3D model / fallbacks). Anchoring on those would target an element that
    /// does not exist yet (raw-set would match nothing) — the caller falls back
    /// to append for them.
    /// </summary>
    private static bool IsStableZOrderAnchor(DocumentFormat.OpenXml.OpenXmlElement el)
    {
        switch (el.LocalName)
        {
            case "sp":
            case "cxnSp":
            case "grpSp":
                return true;
            case "pic":
                // Media (video/audio) is a <p:pic> emitted in a late pass.
                var picXml = el.OuterXml;
                return !picXml.Contains("videoFile", StringComparison.Ordinal)
                    && !picXml.Contains("audioFile", StringComparison.Ordinal);
            case "graphicFrame":
                var gfXml = el.OuterXml;
                // SmartArt (dgm:relIds) and OLE (<p:oleObj>) are emitted late;
                // chart / table graphicFrames are main-walk (stable).
                return !gfXml.Contains("/diagram\"", StringComparison.Ordinal)
                    && !gfXml.Contains(":relIds", StringComparison.Ordinal)
                    && !gfXml.Contains("oleObj", StringComparison.Ordinal);
            default:
                // AlternateContent (3D model / fallbacks) and anything else.
                return false;
        }
    }

    /// <summary>
    /// Resolve a SmartArt sub-part's zip-URI for raw-set targeting. Given a
    /// slide index and a rId (data/layout/colors/quickStyle), returns
    /// e.g. "/ppt/diagrams/data1.xml". Returns null if the rId does not
    /// resolve to a known diagram part type.
    /// </summary>
    internal string? GetSmartArtPartUri(int slideIdx, string relId)
    {
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        try
        {
            var part = slidePart.GetPartById(relId);
            if (part is DiagramDataPart or DiagramLayoutDefinitionPart
                or DiagramColorsPart or DiagramStylePart)
            {
                return part.Uri.OriginalString;
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Per-slide video/audio info for PptxBatchEmitter Phase 3c-media
    /// passthrough. Returns one entry per &lt;p:pic&gt; on the slide whose
    /// nvPr carries &lt;a:videoFile&gt; or &lt;a:audioFile&gt;. Each entry
    /// includes the &lt;p:pic&gt; XML verbatim plus the source's three rIds
    /// (link/media/thumbnail) and the underlying binary streams, so the
    /// emitter can issue an `add-part video` (or `audio`) + a `raw-set`
    /// append on /p:sld/p:cSld/p:spTree that round-trips byte-equal.
    /// </summary>
    internal readonly record struct MediaInfo(
        string PicXml,
        bool IsVideo,
        string LinkRelId,
        string MediaEmbedRelId,
        string ThumbnailRelId,
        byte[] MediaBytes,
        string MediaContentType,
        string MediaExtension,
        byte[] ThumbnailBytes,
        string ThumbnailContentType);

    internal IReadOnlyList<MediaInfo> GetMediaOnSlide(int slideIdx)
    {
        var result = new List<MediaInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];
        var slide = GetSlide(slidePart);
        var spTree = slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return result;

        foreach (var pic in spTree.Descendants<Picture>())
        {
            var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            if (nvPr == null) continue;

            var videoFile = nvPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.VideoFromFile>();
            var audioFile = nvPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.AudioFromFile>();
            bool isVideo = videoFile != null;
            bool isAudio = audioFile != null;
            if (!isVideo && !isAudio) continue;

            string? linkRid = isVideo ? videoFile?.Link?.Value : audioFile?.Link?.Value;
            if (string.IsNullOrEmpty(linkRid)) continue;

            // Locate the p14:media extension carrying the MediaReference rId.
            string? mediaEmbedRid = null;
            var p14Media = nvPr.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
            if (p14Media?.Embed?.Value != null) mediaEmbedRid = p14Media.Embed.Value;
            if (string.IsNullOrEmpty(mediaEmbedRid)) continue;

            // Thumbnail rId from blipFill.
            var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
            var thumbRid = blip?.Embed?.Value;
            if (string.IsNullOrEmpty(thumbRid)) continue;

            // Resolve media binary via either rId. Both VideoReference and
            // MediaReference point at the same MediaDataPart.
            byte[]? mediaBytes = null;
            string? mediaCT = null;
            string? mediaExt = null;
            try
            {
                MediaDataPart? mdp = null;
                foreach (var rel in slidePart.DataPartReferenceRelationships)
                {
                    if (rel.Id == linkRid || rel.Id == mediaEmbedRid)
                    {
                        if (rel.DataPart is MediaDataPart mdp2) { mdp = mdp2; break; }
                    }
                }
                if (mdp != null)
                {
                    using var s = mdp.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    mediaBytes = ms.ToArray();
                    mediaCT = mdp.ContentType;
                    // Extract extension from the part Uri.
                    var u = mdp.Uri.OriginalString;
                    var dot = u.LastIndexOf('.');
                    mediaExt = dot > 0 ? u[dot..] : (isVideo ? ".mp4" : ".mp3");
                }
            }
            catch { }
            if (mediaBytes == null || mediaCT == null || mediaExt == null) continue;

            // Resolve thumbnail binary.
            byte[]? thumbBytes = null;
            string? thumbCT = null;
            try
            {
                var p = slidePart.GetPartById(thumbRid);
                if (p is ImagePart ip)
                {
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    thumbBytes = ms.ToArray();
                    thumbCT = ip.ContentType;
                }
            }
            catch { }
            if (thumbBytes == null || thumbCT == null) continue;

            result.Add(new MediaInfo(
                PicXml: pic.OuterXml,
                IsVideo: isVideo,
                LinkRelId: linkRid,
                MediaEmbedRelId: mediaEmbedRid,
                ThumbnailRelId: thumbRid,
                MediaBytes: mediaBytes,
                MediaContentType: mediaCT,
                MediaExtension: mediaExt,
                ThumbnailBytes: thumbBytes,
                ThumbnailContentType: thumbCT));
        }
        return result;
    }

    /// <summary>
    /// A <c>&lt;p:pic&gt;</c> that hosts a <c>&lt;a:videoFile&gt;</c> /
    /// <c>&lt;a:audioFile&gt;</c> but is NOT a modern embedded-media shape
    /// (no <c>p14:media</c> extension resolving to a local MediaDataPart) —
    /// typically a legacy PowerPoint 2007 external linked movie
    /// (<c>videoFile r:link</c> pointing at a <c>TargetMode="External"</c>
    /// file:// URI, with a local poster image in the blipFill). GetMediaOnSlide
    /// rejects these (it requires an embedded MediaDataPart), and the typed
    /// walk skips <c>video</c>/<c>audio</c> nodes, so without this pass the
    /// entire picture is silently dropped on dump∘replay. We round-trip the
    /// shape verbatim plus its external link relationship(s) and poster
    /// image(s), so the poster still renders and the videoFile r:link no
    /// longer dangles.
    /// </summary>
    internal readonly record struct ExternalMediaPicInfo(
        string PicXml,
        IReadOnlyList<(string Rid, string RelType, string Target)> ExternalRels,
        IReadOnlyList<string> ImageRids);

    internal IReadOnlyList<ExternalMediaPicInfo> GetExternalMediaPicsOnSlide(int slideIdx)
    {
        var result = new List<ExternalMediaPicInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];
        var slide = GetSlide(slidePart);
        var spTree = slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return result;

        foreach (var pic in spTree.Descendants<Picture>())
        {
            var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            if (nvPr == null) continue;

            var videoFile = nvPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.VideoFromFile>();
            var audioFile = nvPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.AudioFromFile>();
            if (videoFile == null && audioFile == null) continue;

            // If a p14:media extension resolves to a local MediaDataPart, this
            // is a modern embedded-media shape owned by EmitMediaForSlide; skip
            // it here to avoid double-emit.
            var p14Media = nvPr.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
            var mediaEmbedRid = p14Media?.Embed?.Value;
            bool hasLocalMedia = false;
            if (!string.IsNullOrEmpty(mediaEmbedRid))
            {
                foreach (var rel in slidePart.DataPartReferenceRelationships)
                    if (rel.Id == mediaEmbedRid && rel.DataPart is MediaDataPart) { hasLocalMedia = true; break; }
            }
            if (hasLocalMedia) continue;

            // Collect every relationship id the pic references (r:embed / r:link /
            // r:id on any descendant attribute), then classify each as an
            // external relationship (carry via add-part extrel) or a local
            // ImagePart (carry via add-part image).
            var refRids = new HashSet<string>(StringComparer.Ordinal);
            foreach (Match m in Regex.Matches(pic.OuterXml, @"r:(?:embed|link|id)=""(rId\d+)"""))
                refRids.Add(m.Groups[1].Value);

            var extRels = new List<(string, string, string)>();
            var imageRids = new List<string>();
            foreach (var rid in refRids)
            {
                var ext = slidePart.ExternalRelationships.FirstOrDefault(r => r.Id == rid);
                if (ext != null)
                {
                    extRels.Add((rid, ext.RelationshipType, ext.Uri.OriginalString));
                    continue;
                }
                try
                {
                    if (slidePart.GetPartById(rid) is ImagePart) imageRids.Add(rid);
                }
                catch { /* media/other data-part rel — not an ImagePart, ignore */ }
            }

            result.Add(new ExternalMediaPicInfo(pic.OuterXml, extRels, imageRids));
        }
        return result;
    }

    /// <summary>
    /// chartEx (cx: extension chart) parts referenced from a slide by rId —
    /// payload for the emitter's `add-part chartex` carrier. Each entry is the
    /// chartEx XML plus its typed children (colors / style sidecars, embedded
    /// xlsx workbook) with their rIds, all base64.
    /// </summary>
    internal readonly record struct ChartExInfo(
        string RelId,
        string XmlBase64,
        string? ColorsRelId, string? ColorsBase64,
        string? StyleRelId, string? StyleBase64,
        string? PackageRelId, string? PackageBase64, string? PackageContentType);

    internal IReadOnlyList<ChartExInfo> GetChartExPartsByRelId(int slideIdx, IEnumerable<string> rids)
    {
        var result = new List<ChartExInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];

        static string ReadB64(OpenXmlPart p)
        {
            using var s = p.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return Convert.ToBase64String(ms.ToArray());
        }

        foreach (var rid in rids)
        {
            OpenXmlPart? part = null;
            try { part = slidePart.GetPartById(rid); } catch { }
            if (part is not ExtendedChartPart cxPart) continue;

            string? colorsRid = null, colorsB64 = null, styleRid = null, styleB64 = null;
            string? pkgRid = null, pkgB64 = null, pkgCT = null;
            foreach (var pair in cxPart.Parts)
            {
                switch (pair.OpenXmlPart)
                {
                    case ChartColorStylePart ccs:
                        colorsRid = pair.RelationshipId; colorsB64 = ReadB64(ccs); break;
                    case ChartStylePart cst:
                        styleRid = pair.RelationshipId; styleB64 = ReadB64(cst); break;
                    case EmbeddedPackagePart epp:
                        pkgRid = pair.RelationshipId; pkgB64 = ReadB64(epp); pkgCT = epp.ContentType; break;
                }
            }
            result.Add(new ChartExInfo(rid, ReadB64(cxPart),
                colorsRid, colorsB64, styleRid, styleB64, pkgRid, pkgB64, pkgCT));
        }
        return result;
    }

    /// <summary>
    /// Per-slide am3d 3D-model info for PptxBatchEmitter Phase 3c-3d
    /// passthrough. Returns one entry per &lt;mc:AlternateContent&gt; block
    /// whose &lt;mc:Choice Requires="am3d"&gt; carries an &lt;am3d:model3d&gt;
    /// element. Each entry includes the AlternateContent XML verbatim plus
    /// the source's two rIds (the model3d ExtendedPart and the shared
    /// thumbnail ImagePart) and the underlying binary streams, so the
    /// emitter can issue an `add-part model3d` + a `raw-set` append on
    /// /p:sld/p:cSld/p:spTree that round-trips byte-equal.
    /// </summary>
    internal readonly record struct Model3dInfo(
        string AlternateContentXml,
        string Model3dRelId,
        string ThumbnailRelId,
        byte[] Model3dBytes,
        string Model3dContentType,
        string Model3dExtension,
        byte[] ThumbnailBytes,
        string ThumbnailContentType);

    private static string InjectAmbientXmlnsOnRoot(string sliceXml,
        (string Prefix, string Uri)[] decls)
    {
        if (string.IsNullOrEmpty(sliceXml) || sliceXml[0] != '<') return sliceXml;
        int gt = sliceXml.IndexOf('>');
        if (gt <= 0) return sliceXml;
        var head = sliceXml[..gt];
        var tail = sliceXml[gt..];
        // Skip injecting decls that the root already carries (avoids
        // duplicate xmlns attributes which is illegal).
        var sb = new System.Text.StringBuilder(head);
        foreach (var (prefix, uri) in decls)
        {
            if (head.Contains($"xmlns:{prefix}=\"", StringComparison.Ordinal)) continue;
            // Only inject if the slice actually references this prefix.
            if (!sliceXml.Contains($"<{prefix}:", StringComparison.Ordinal)
                && !sliceXml.Contains($" {prefix}:", StringComparison.Ordinal))
                continue;
            sb.Append(" xmlns:").Append(prefix).Append("=\"").Append(uri).Append('"');
        }
        sb.Append(tail);
        return sb.ToString();
    }

    internal IReadOnlyList<Model3dInfo> GetModel3dOnSlide(int slideIdx)
    {
        var result = new List<Model3dInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];

        // Read raw slide XML — am3d lives under <mc:AlternateContent> which
        // the typed SDK tree exposes only as OpenXmlUnknownElement. Walking
        // the raw stream avoids the awkward typed traversal and gives us a
        // straight slice for raw-set passthrough.
        string slideXml;
        using (var s = slidePart.GetStream())
        using (var sr = new StreamReader(s))
            slideXml = sr.ReadToEnd();

        // Parse the slide XML and walk for <mc:AlternateContent> elements
        // whose Choice has Requires="am3d". We extract slices by element
        // identity rather than textual regex because the SDK may re-prefix
        // the relationships namespace (e.g. p10:embed instead of r:embed)
        // when re-serialising an unknown subtree on round-trip — text-only
        // matching misses these.
        System.Xml.Linq.XDocument slideDoc;
        try { slideDoc = System.Xml.Linq.XDocument.Parse(slideXml); }
        catch { return result; }
        System.Xml.Linq.XNamespace mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        System.Xml.Linq.XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        System.Xml.Linq.XNamespace am3dNs = "http://schemas.microsoft.com/office/drawing/2017/model3d";
        System.Xml.Linq.XNamespace aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        foreach (var ac in slideDoc.Descendants(mcNs + "AlternateContent").ToList())
        {
            var choice = ac.Element(mcNs + "Choice");
            var requires = choice?.Attribute("Requires")?.Value;
            if (choice == null || requires == null || !requires.Contains("am3d", StringComparison.Ordinal)) continue;

            var model3d = choice.Descendants(am3dNs + "model3d").FirstOrDefault();
            var m3dRidAttr = model3d?.Attribute(rNs + "embed");
            if (m3dRidAttr == null) continue;
            var m3dRid = m3dRidAttr.Value;

            // Thumbnail rId: prefer <am3d:blip r:embed=…> in <am3d:raster>;
            // fall back to <a:blip r:embed=…> in the Fallback <p:pic>.
            string? thumbRid = model3d!.Descendants(am3dNs + "blip")
                .Select(b => b.Attribute(rNs + "embed")?.Value)
                .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            if (string.IsNullOrEmpty(thumbRid))
            {
                thumbRid = ac.Descendants(aNs + "blip")
                    .Select(b => b.Attribute(rNs + "embed")?.Value)
                    .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            }
            if (string.IsNullOrEmpty(thumbRid)) continue;

            // Re-serialise the AlternateContent subtree as the slice. This
            // gives a canonical XML form that NormalizeSlideRawSlice can
            // further harmonise — both round-1 (read from source) and
            // round-2 (read from B.pptx) paths funnel through XLinq here,
            // so namespace prefix drift (p10:embed vs r:embed) reconciles.
            var slice = ac.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

            // Resolve the model3d binary via the slide's ExtendedParts
            // relationship dictionary. AddExtendedPart routes to
            // ExtendedParts, NOT to a typed part of slidePart.Parts; we
            // must reach in by rel id.
            byte[]? m3dBytes = null;
            string? m3dCT = null;
            string m3dExt = ".glb";
            try
            {
                // Extended parts (created via AddExtendedPart) live in
                // slidePart.Parts under their pinned rel id; GetPartById
                // resolves any internal child part by relationship id
                // regardless of typing.
                var part = slidePart.GetPartById(m3dRid);
                if (part != null)
                {
                    using var s = part.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    m3dBytes = ms.ToArray();
                    m3dCT = part.ContentType;
                    var u = part.Uri.OriginalString;
                    var dot = u.LastIndexOf('.');
                    if (dot > 0) m3dExt = u[dot..];
                }
            }
            catch { }
            if (m3dBytes == null || string.IsNullOrEmpty(m3dCT)) continue;

            // Resolve thumbnail bytes from the shared ImagePart.
            byte[]? thumbBytes = null;
            string? thumbCT = null;
            try
            {
                var tp = slidePart.GetPartById(thumbRid);
                if (tp is ImagePart ip)
                {
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    thumbBytes = ms.ToArray();
                    thumbCT = ip.ContentType;
                }
            }
            catch { }
            if (thumbBytes == null || string.IsNullOrEmpty(thumbCT)) continue;

            result.Add(new Model3dInfo(
                AlternateContentXml: slice,
                Model3dRelId: m3dRid,
                ThumbnailRelId: thumbRid,
                Model3dBytes: m3dBytes,
                Model3dContentType: m3dCT!,
                Model3dExtension: m3dExt,
                ThumbnailBytes: thumbBytes,
                ThumbnailContentType: thumbCT!));
        }
        return result;
    }

    /// <summary>
    /// Per-slide OLE-embed info for PptxBatchEmitter Phase 3c-ole passthrough.
    /// Returns one entry per &lt;p:graphicFrame&gt; whose
    /// &lt;a:graphicData uri="…/presentationml/2006/ole"&gt; carries a
    /// &lt;p:oleObj&gt; element. Each entry includes the graphicFrame XML
    /// verbatim plus the source's two rIds (the OLE part and the icon
    /// thumbnail ImagePart) and the underlying binary streams, so the
    /// emitter can issue an `add-part ole` + a `raw-set` append on
    /// /p:sld/p:cSld/p:spTree that round-trips byte-equal.
    /// </summary>
    internal readonly record struct OleInfo(
        string GraphicFrameXml,
        string OleRelId,
        string ThumbnailRelId,
        byte[] OleBytes,
        string OleContentType,
        string OleExtension,
        byte[] ThumbnailBytes,
        string ThumbnailContentType,
        // Linked (TargetMode=External) OLE: no embedded payload; the emitter
        // recreates the external rel instead of an add-part ole payload.
        string? LinkedTarget = null,
        string? LinkedRelType = null);

    internal IReadOnlyList<OleInfo> GetOlesOnSlide(int slideIdx)
    {
        var result = new List<OleInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];

        // Read raw slide XML — the OleObject child of GraphicData is
        // typed in the SDK (DocumentFormat.OpenXml.Presentation.OleObject)
        // but the whole graphicFrame slice is what we want to raw-set, so
        // a textual XLinq walk is cleaner. Matches the model3d Phase 3c-3d
        // approach: read raw, parse, slice by element identity (no regex)
        // to dodge SDK namespace-prefix drift on round 2.
        string slideXml;
        using (var s = slidePart.GetStream())
        using (var sr = new StreamReader(s))
            slideXml = sr.ReadToEnd();

        System.Xml.Linq.XDocument slideDoc;
        try { slideDoc = System.Xml.Linq.XDocument.Parse(slideXml); }
        catch { return result; }
        System.Xml.Linq.XNamespace pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        System.Xml.Linq.XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        System.Xml.Linq.XNamespace aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string oleUri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

        foreach (var gf in slideDoc.Descendants(pNs + "graphicFrame").ToList())
        {
            var gd = gf.Descendants(aNs + "graphicData").FirstOrDefault();
            var uri = gd?.Attribute("uri")?.Value;
            if (uri == null || !uri.Equals(oleUri, StringComparison.Ordinal)) continue;

            // The oleObj may sit directly under graphicData OR wrapped in
            // <mc:AlternateContent><mc:Choice Requires="v">…</mc:Choice>
            // <mc:Fallback>…</mc:Fallback> (legacy VML-annotated OLE,
            // graphic-stroke). Choice and Fallback share the same r:id, so
            // any descendant works for rel resolution; the slice carries the
            // whole wrapper verbatim either way.
            // Prefer the oleObj that carries the thumbnail <p:pic> (the
            // mc:Fallback one) — the mc:Choice twin holds only <p:link/> and
            // would trip the no-thumbnail legacy skip below.
            var oleObj = gd!.Element(pNs + "oleObj")
                ?? gd.Descendants(pNs + "oleObj")
                     .OrderByDescending(o => o.Descendants(aNs + "blip").Any())
                     .FirstOrDefault();
            if (oleObj == null) continue;
            var oleRidAttr = oleObj.Attribute(rNs + "id");
            if (oleRidAttr == null) continue;
            var oleRid = oleRidAttr.Value;

            // Thumbnail icon: <p:pic>/<p:blipFill>/<a:blip r:embed="…"/>
            // inside the <p:oleObj>. Modern DrawingML OLE embeds carry one.
            // A legacy OLE (<p:oleObj spid="_x0000_s…"><p:embed/></p:oleObj>
            // whose cached visual is a VML shape referenced by spid) has NO
            // inner pic blip. Skip it: round-tripping just the graphicFrame +
            // payload (without the VML drawing part that backs the spid) yields
            // a deck real PowerPoint refuses to open — worse than dropping the
            // object. Full legacy-VML-OLE round-trip is a deferred feature.
            var thumbRid = oleObj.Descendants(aNs + "blip")
                .Select(b => b.Attribute(rNs + "embed")?.Value)
                .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            if (string.IsNullOrEmpty(thumbRid)) continue;

            // Re-serialise the graphicFrame subtree as the slice. Funnels
            // both rounds through XLinq for prefix-drift reconciliation.
            var slice = gf.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

            // Resolve the OLE payload. Could be EmbeddedPackagePart (modern
            // OOXML container) or EmbeddedObjectPart (generic binary / .bin
            // OLE10 stream / legacy .doc/.xls). GetPartById resolves either.
            byte[]? oleBytes = null;
            string? oleCT = null;
            string oleExt = ".bin";
            try
            {
                var part = slidePart.GetPartById(oleRid);
                if (part != null)
                {
                    using var s = part.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    oleBytes = ms.ToArray();
                    oleCT = part.ContentType;
                    var u = part.Uri.OriginalString;
                    var dot = u.LastIndexOf('.');
                    if (dot > 0) oleExt = u[dot..];
                }
            }
            catch { }
            if (oleBytes == null || string.IsNullOrEmpty(oleCT))
            {
                // LINKED OLE: the r:id is a TargetMode="External" relationship
                // (file:// link with <p:link/> inside p:oleObj) — there is no
                // embedded payload part to carry. Round-trip the graphicFrame
                // verbatim plus the external rel and the thumbnail image, or
                // the whole object silently vanishes (graphic-stroke).
                var extRel = slidePart.ExternalRelationships.FirstOrDefault(r => r.Id == oleRid);
                if (extRel != null)
                {
                    byte[]? linkThumbBytes = null;
                    string? linkThumbCT = null;
                    try
                    {
                        if (slidePart.GetPartById(thumbRid!) is ImagePart lip)
                        {
                            using var ls = lip.GetStream();
                            using var lms = new MemoryStream();
                            ls.CopyTo(lms);
                            linkThumbBytes = lms.ToArray();
                            linkThumbCT = lip.ContentType;
                        }
                    }
                    catch { }
                    if (linkThumbBytes != null && linkThumbCT != null)
                    {
                        result.Add(new OleInfo(
                            GraphicFrameXml: slice,
                            OleRelId: oleRid,
                            ThumbnailRelId: thumbRid!,
                            OleBytes: Array.Empty<byte>(),
                            OleContentType: "",
                            OleExtension: "",
                            ThumbnailBytes: linkThumbBytes,
                            ThumbnailContentType: linkThumbCT,
                            LinkedTarget: extRel.Uri.OriginalString,
                            LinkedRelType: extRel.RelationshipType));
                    }
                }
                continue;
            }

            // Resolve the thumbnail icon ImagePart bytes.
            byte[]? thumbBytes = null;
            string? thumbCT = null;
            try
            {
                var tp = slidePart.GetPartById(thumbRid);
                if (tp is ImagePart ip)
                {
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    thumbBytes = ms.ToArray();
                    thumbCT = ip.ContentType;
                }
            }
            catch { }
            if (thumbBytes == null || string.IsNullOrEmpty(thumbCT)) continue;

            result.Add(new OleInfo(
                GraphicFrameXml: slice,
                OleRelId: oleRid,
                ThumbnailRelId: thumbRid,
                OleBytes: oleBytes,
                OleContentType: oleCT!,
                OleExtension: oleExt,
                ThumbnailBytes: thumbBytes,
                ThumbnailContentType: thumbCT!));
        }
        return result;
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
        var blip = ResolvePictureBlip(pic);
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

    // Resolve a picture's effective fill <a:blip>, transparently unwrapping an
    // <mc:AlternateContent> wrapper. Mac PowerPoint emits the blipFill inside
    // AlternateContent — a Requires="ma" <mc:Choice> holding a Mac-only source
    // (often an image PDF) plus an <mc:Fallback> holding the standards-
    // compliant raster (PNG/JPEG) that every other renderer uses. When the
    // blipFill is NOT a direct child of <p:pic>, `pic.BlipFill` is null and the
    // picture would otherwise read as image-less and be DROPPED on dump→replay
    // (silent element loss). Prefer the Fallback blip (what Windows/real Office
    // renders); fall back to the Choice, then any descendant blip.
    private static DocumentFormat.OpenXml.Drawing.Blip? ResolvePictureBlip(Picture pic)
    {
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        if (blip != null) return blip;

        var altContent = pic.GetFirstChild<DocumentFormat.OpenXml.AlternateContent>();
        if (altContent != null)
        {
            var fallback = altContent
                .GetFirstChild<DocumentFormat.OpenXml.AlternateContentFallback>();
            blip = fallback?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                .FirstOrDefault();
            if (blip != null) return blip;

            var choice = altContent
                .GetFirstChild<DocumentFormat.OpenXml.AlternateContentChoice>();
            blip = choice?.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                .FirstOrDefault();
            if (blip != null) return blip;
        }

        // Last resort: any nested blip (covers other MC nestings). Within a
        // <p:pic> the only blips are fill blips, so first-in-document-order is
        // the effective source.
        return pic.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
    }

    // Return the verbatim <a:clrChange> outer XML on a picture's <a:blip>,
    // or null when absent. Used by PptxBatchEmitter.EmitPicture to round-
    // trip color-change adjustments (recolor) — there is no typed Set
    // vocabulary for clrChange today, so the emitter copies the element
    // through a raw-set passthrough on the freshly-added picture's blip.
    // Mirrors the GetImageBinary path-resolution preamble verbatim.
    public string? GetPictureBlipClrChangeXml(string picturePath)
    {
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
        Picture? pic = byId
            ? pictures.FirstOrDefault(p =>
                p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value
                    == (uint)idOrIdx)
            : (idOrIdx >= 1 && idOrIdx <= pictures.Count ? pictures[idOrIdx - 1] : null);
        if (pic == null) return null;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        var clrChange = blip?.GetFirstChild<DocumentFormat.OpenXml.Drawing.ColorChange>();
        return clrChange?.OuterXml;
    }

    // Picture-in-placeholder round-trip. A picture that fills a layout
    // placeholder carries `<p:nvPr><p:ph type="pic" idx="N"/></p:nvPr>` and an
    // empty (xfrm-less) `<p:spPr/>`, inheriting its position+size from the
    // layout placeholder. EmitPicture rebuilds it as a free-floating picture:
    // the `<p:ph>` is dropped and AddPicture stamps a default xfrm, so it
    // renders at the wrong size/offset. Capture the `<p:ph>` element and — only
    // when the source has no explicit `<a:xfrm>` — the source `<p:spPr>`, so
    // EmitPicture can re-inject them via raw-set and let the layout drive the
    // geometry again. Mirrors the path-resolution preamble in
    // GetPictureBlipClrChangeXml verbatim.
    public (string? PhXml, string? InheritSpPrXml) GetPicturePlaceholderRoundtripXml(string picturePath)
    {
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return (null, null);
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return (null, null);
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return (null, null);
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = byId
            ? pictures.FirstOrDefault(p =>
                p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value
                    == (uint)idOrIdx)
            : (idOrIdx >= 1 && idOrIdx <= pictures.Count ? pictures[idOrIdx - 1] : null);
        if (pic == null) return (null, null);

        var ph = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<PlaceholderShape>();
        if (ph == null) return (null, null);

        var spPr = pic.ShapeProperties;
        var hasXfrm = spPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Transform2D>() != null;
        // Restore the source spPr (dropping AddPicture's default xfrm) only when
        // the source inherited its geometry. If the source pinned an explicit
        // xfrm, the x/y/width/height props already round-trip it.
        string? inheritSpPr = hasXfrm ? null : (spPr?.OuterXml);
        return (ph.OuterXml, inheritSpPr);
    }

    // R56 bt-6: return outer XML for every <a:blip> child the typed Add/Set
    // surface does NOT cover, so dump→batch can re-inject them via raw-set
    // passthrough. Standard typed-handled children (filtered out): a:alphaModFix
    // (opacity), a:biLevel, a:duotone, a:lum + legacy lumOff/lumMod
    // (brightness/contrast), a:clrChange (already round-tripped via the
    // dedicated GetPictureBlipClrChangeXml + EmitPicture raw-set).
    // Everything else — alphaBiLevel, alphaCeiling, alphaFloor, alphaInv,
    // alphaMod, alphaRepl, blur, clrRepl, fillOverlay, grayscl, hsl, tint,
    // extLst, plus non-schema extension elements like <a:colorMod> seen in
    // the wild — was silently dropped on dump. Mirrors the path-resolution
    // preamble in GetPictureBlipClrChangeXml verbatim.
    public IReadOnlyList<string> GetPictureBlipPassthroughChildrenXml(string picturePath)
    {
        var empty = (IReadOnlyList<string>)Array.Empty<string>();
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return empty;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return empty;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return empty;
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = byId
            ? pictures.FirstOrDefault(p =>
                p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value
                    == (uint)idOrIdx)
            : (idOrIdx >= 1 && idOrIdx <= pictures.Count ? pictures[idOrIdx - 1] : null);
        if (pic == null) return empty;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        if (blip == null) return empty;
        var result = new List<string>();
        foreach (var kid in blip.ChildElements)
        {
            // Skip the typed-handled blip children — Add/Set already
            // round-trips these via Format keys.
            if (kid is DocumentFormat.OpenXml.Drawing.AlphaModulationFixed) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.BiLevel) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.Duotone) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.LuminanceEffect) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.ColorChange) continue;
            // Legacy invalid markup written by older builds — NodeBuilder
            // already maps these to brightness/contrast Format keys so the
            // Set-side path re-writes them as <a:lum>. Don't double-emit.
            if (kid.LocalName is "lumOff" or "lumMod"
                && kid.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main")
                continue;
            result.Add(kid.OuterXml);
        }
        return result;
    }

    /// <summary>
    /// Companion binary parts a picture's <a:blip> references from inside its
    /// <a:extLst> — beyond the main <c>r:embed</c>. The common cases:
    /// <list type="bullet">
    /// <item>HD Photo backup layer (<c>&lt;a14:imgProps&gt;&lt;a14:imgLayer r:embed&gt;</c>
    /// → a <c>.wdp</c> part, relationship type <c>.../2007/relationships/hdphoto</c>)
    /// carrying advanced image effects.</item>
    /// <item>SVG companion (<c>&lt;asvg:svgBlip r:embed&gt;</c> → the vector
    /// original behind a raster fallback).</item>
    /// </list>
    /// The main image round-trips via <c>add picture</c> and the blip's extLst is
    /// re-appended verbatim by <see cref="GetPictureBlipPassthroughChildrenXml"/>,
    /// so the companion <c>r:embed="rIdN"</c> survives — but the part it points at
    /// was never re-emitted, leaving a dangling relationship (lost image-effects
    /// layer; stricter consumers reject the package). Surfaced as
    /// (rId, relationship-type, content-type, target-ext, base64) so EmitPicture
    /// can pin each via an <c>add-part extpart</c> row. Mirrors the master/layout
    /// image carrier. Returns empty when the blip carries no companion references.
    /// </summary>
    public readonly record struct BlipCompanionInfo(
        string RelId, string RelType, string ContentType, string TargetExt, string Base64Data);

    public IReadOnlyList<BlipCompanionInfo> GetPictureBlipCompanionParts(string picturePath)
    {
        var empty = (IReadOnlyList<BlipCompanionInfo>)Array.Empty<BlipCompanionInfo>();
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return empty;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return empty;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return empty;
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = byId
            ? pictures.FirstOrDefault(p =>
                p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value
                    == (uint)idOrIdx)
            : (idOrIdx >= 1 && idOrIdx <= pictures.Count ? pictures[idOrIdx - 1] : null);
        if (pic == null) return empty;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        if (blip == null) return empty;
        var mainEmbed = blip.Embed?.Value;

        const string relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var result = new List<BlipCompanionInfo>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        // Walk every descendant of the blip (the extLst lives there) and collect
        // r:embed / r:link / r:id references other than the main image.
        foreach (var el in blip.Descendants())
        {
            foreach (var attr in el.GetAttributes())
            {
                if (attr.NamespaceUri != relNs) continue;
                if (attr.LocalName is not ("embed" or "link" or "id")) continue;
                var rid = attr.Value;
                if (string.IsNullOrEmpty(rid) || rid == mainEmbed || !seen.Add(rid)) continue;
                try
                {
                    var part = slidePart.GetPartById(rid);
                    using var s = part.GetStream(FileMode.Open, FileAccess.Read);
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    var ext = System.IO.Path.GetExtension(part.Uri.OriginalString);
                    if (string.IsNullOrEmpty(ext)) ext = ".bin";
                    result.Add(new BlipCompanionInfo(
                        rid, part.RelationshipType, part.ContentType, ext,
                        Convert.ToBase64String(ms.ToArray())));
                }
                catch { /* external / unresolvable — skip (dangling already) */ }
            }
        }
        return result;
    }

    // Probe whether a shape's NonVisualDrawingProperties carries a
    // hlinkClick child. Used by PptxBatchEmitter.EmitShape to disambiguate
    // a Format["link"] surfaced by NodeBuilder's single-run shortcut
    // (run-level hlinkClick promoted onto the shape Format bag for Get
    // convenience) from a true shape-level hlinkClick. The dump path
    // should emit shape-level link= on Add only when the link is truly on
    // the cNvPr — otherwise the run-level emit duplicates and AddShape
    // fabricates a shape-level hyperlink that the source never had.
    internal bool ShapeHasCNvPrHyperlink(string shapePath)
        => GetShapeCNvPrHyperlinkInfo(shapePath).HasShapeLink;

    // CONSISTENCY(shape-link-source-readback): when a shape carries BOTH a
    // cNvPr.hlinkClick AND a first-run rPr.hlinkClick, NodeBuilder promotes
    // the RUN url onto Format["link"] (first-run wins; line ~960). The dump
    // emitter needs the actual shape-level url so it can emit `add shape
    // link=<shape-url>` instead of inheriting the run url. Return the
    // shape-level url + tooltip (or null/empty) alongside the boolean.
    internal (bool HasShapeLink, string? Url, string? Tooltip) GetShapeCNvPrHyperlinkInfo(string shapePath)
    {
        // Accept positional /shape[N] and @id= forms. NodeBuilder emits the
        // @id= form for shapes with a known cNvPr.Id (the typical case);
        // the typed dump walk passes shapeNode.Path verbatim.
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/(?:shape|textbox|title|equation|placeholder)\[(@id=)?(\d+)\]$");
        if (!m.Success) return (false, null, null);
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        var byId = m.Groups[3].Value.Length > 0;
        var shapeIdx = int.Parse(m.Groups[4].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return (false, null, null);
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return (false, null, null);
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(\d+)\]"))
        {
            var gIdx = int.Parse(gm.Groups[1].Value);
            var groupsHere = scope.Elements<GroupShape>().ToList();
            if (gIdx < 1 || gIdx > groupsHere.Count) return (false, null, null);
            scope = groupsHere[gIdx - 1];
        }
        Shape? shape;
        if (byId)
        {
            shape = scope.Elements<Shape>().FirstOrDefault(
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)shapeIdx);
            if (shape == null) return (false, null, null);
        }
        else
        {
            var shapes = scope.Elements<Shape>().ToList();
            if (shapeIdx < 1 || shapeIdx > shapes.Count) return (false, null, null);
            shape = shapes[shapeIdx - 1];
        }
        var nvDp = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
        var hlClick = nvDp?.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>();
        if (hlClick == null) return (false, null, null);
        var url = ReadHyperlinkOnClickUrl(hlClick, slidePart);
        var tip = hlClick.Tooltip?.Value;
        return (true, url, tip);
    }

    // Return the verbatim <p:style> outer XML on a shape (the lnRef / fillRef
    // / effectRef / fontRef theme-reference block) plus the shape's ordinal
    // among <p:sp> siblings in its parent shapeTree/group, or null when
    // either the shape has no <p:style> or the path doesn't resolve. Used by
    // PptxBatchEmitter.EmitShape to round-trip the style reference block
    // through a raw-set passthrough — there is no typed Add/Set vocabulary
    // for these theme-style refs today and the source block was silently
    // dropped on dump->replay before this hook.
    // CONSISTENCY: mirrors GetShapeCNvPrHyperlinkInfo's path-resolution
    // preamble (group-chain + @id= / positional shape index).
    // expectId: when the caller knows the source node's cNvPr id, verify the
    // path-resolved shape IS that shape. A node emitted with a positional path
    // can resolve to a DIFFERENT sp at that ordinal and steal its <p:style> —
    // bnc889755's sldNum placeholder duplicated TextBox 10's style onto the
    // same replayed sp, and the doubled <p:style> child made PowerPoint refuse
    // the deck (0x80070570).
    internal (string Xml, int SpOrdinal)? GetShapeStyleXmlWithOrdinal(string shapePath, uint? expectId = null)
    {
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]((?:/group\[(?:@id=)?\d+\])*)/(shape|textbox|title|equation|placeholder)\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        // Query emits PLACEHOLDER paths with a placeholder-scoped positional
        // index (/slide[15]/placeholder[1] = the slide's FIRST placeholder,
        // not its first <p:sp>). Resolving that index against all shapes
        // picked an unrelated sp — the expectId guard then nulled the probe
        // and the placeholder's <p:style> silently vanished (customGeo).
        var segIsPlaceholder = m.Groups[3].Value == "placeholder";
        var byId = m.Groups[4].Value.Length > 0;
        var shapeIdx = int.Parse(m.Groups[5].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return null;
        // Group ancestors arrive as either /group[K] (positional, the form the
        // batch emitter builds) or /group[@id=N] (the form Query/Get returns);
        // resolve both so group-nested shapes round-trip their <p:style>.
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(@id=)?(\d+)\]"))
        {
            var groupsHere = scope.Elements<GroupShape>().ToList();
            GroupShape? grp;
            if (gm.Groups[1].Value.Length > 0)
                grp = groupsHere.FirstOrDefault(g =>
                    g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                        == (uint)int.Parse(gm.Groups[2].Value));
            else
            {
                var gIdx = int.Parse(gm.Groups[2].Value);
                grp = (gIdx >= 1 && gIdx <= groupsHere.Count) ? groupsHere[gIdx - 1] : null;
            }
            if (grp == null) return null;
            scope = grp;
        }
        var shapes = scope.Elements<Shape>().ToList();
        Shape? shape;
        int ordinal;
        if (byId)
        {
            shape = shapes.FirstOrDefault(
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)shapeIdx);
            if (shape == null) return null;
            ordinal = shapes.IndexOf(shape) + 1;
        }
        else if (segIsPlaceholder)
        {
            var phShapes = shapes.Where(sp2 =>
                sp2.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<PlaceholderShape>() != null).ToList();
            if (shapeIdx < 1 || shapeIdx > phShapes.Count) return null;
            shape = phShapes[shapeIdx - 1];
            ordinal = shapes.IndexOf(shape) + 1;
        }
        else
        {
            if (shapeIdx < 1 || shapeIdx > shapes.Count) return null;
            shape = shapes[shapeIdx - 1];
            ordinal = shapeIdx;
        }
        if (expectId.HasValue
            && shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value != expectId.Value)
            return null;
        var styleEl = shape.GetFirstChild<ShapeStyle>();
        if (styleEl == null) return null;
        return (styleEl.OuterXml, ordinal);
    }

    // Chart-internal image parts (chartSpace/plotArea/series <a:blipFill
    // r:embed> targets). The semantic chart rebuild creates a FRESH ChartPart,
    // so these must be re-attached with pinned rIds or the verbatim spPr
    // fills dangle. Enumerate by the chart's 1-based ordinal on the slide.
    internal IReadOnlyList<MasterImageInfo> GetChartImageParts(int slideIdx, int chartOrdinal)
    {
        var result = new List<MasterImageInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];
        var chartParts = slidePart.ChartParts.ToList();
        if (chartOrdinal < 1 || chartOrdinal > chartParts.Count) return result;
        var cp = chartParts[chartOrdinal - 1];
        foreach (var pair in cp.Parts)
        {
            if (pair.OpenXmlPart is not ImagePart ip) continue;
            using var st = ip.GetStream();
            using var ms = new MemoryStream();
            st.CopyTo(ms);
            result.Add(new MasterImageInfo(pair.RelationshipId, ip.ContentType,
                Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    // Modern chart styling sub-parts (ChartStylePart "style1.xml" /
    // ChartColorStylePart "colors1.xml") attached to a slide's Nth ChartPart.
    // They drive PowerPoint-2013+ gridline tint / palette / effect defaults;
    // the semantic chart rebuild dropped them, flattening the chart to the
    // app default style. Kind is "style" or "colors".
    internal IReadOnlyList<(string Kind, string RelId, string Base64Data)>
        GetChartStyleParts(int slideIdx, int chartOrdinal)
    {
        var result = new List<(string, string, string)>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var chartParts = parts[slideIdx - 1].ChartParts.ToList();
        if (chartOrdinal < 1 || chartOrdinal > chartParts.Count) return result;
        foreach (var pair in chartParts[chartOrdinal - 1].Parts)
        {
            string kind;
            if (pair.OpenXmlPart is ChartStylePart) kind = "style";
            else if (pair.OpenXmlPart is ChartColorStylePart) kind = "colors";
            // Theme override (themeOverride1.xml) — swaps the accent palette
            // for THIS chart only; dropping it recolors every theme-inherited
            // series.
            else if (pair.OpenXmlPart is ThemeOverridePart) kind = "themeoverride";
            else continue;
            using var st = pair.OpenXmlPart.GetStream();
            using var ms = new MemoryStream();
            st.CopyTo(ms);
            result.Add((kind, pair.RelationshipId, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    // Whole-table effects: <a:tblPr><a:effectLst> verbatim + the hosting
    // graphicFrame's 1-based ordinal among the slide's graphicFrames, so the
    // emitter can raw-set PREPEND it into the replayed tblPr (schema places
    // effectLst before tableStyleId). Null when the table has none.
    internal (string Xml, int GfOrdinal)? GetTableEffectsXmlWithOrdinal(string tablePath)
    {
        var m = Regex.Match(tablePath, @"^/slide\[(\d+)\]/table\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var byId = m.Groups[2].Value.Length > 0;
        var tblIdx = int.Parse(m.Groups[3].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var spTree = GetSlide(parts[slideIdx - 1]).CommonSlideData?.ShapeTree;
        if (spTree == null) return null;
        var gfs = spTree.Elements<GraphicFrame>().ToList();
        var tableGfs = gfs.Where(g => g.Descendants<DocumentFormat.OpenXml.Drawing.Table>().Any()).ToList();
        GraphicFrame? gf;
        if (byId)
            gf = tableGfs.FirstOrDefault(g =>
                g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)tblIdx);
        else
            gf = (tblIdx >= 1 && tblIdx <= tableGfs.Count) ? tableGfs[tblIdx - 1] : null;
        if (gf == null) return null;
        var fx = gf.Descendants<DocumentFormat.OpenXml.Drawing.Table>().FirstOrDefault()
            ?.GetFirstChild<DocumentFormat.OpenXml.Drawing.TableProperties>()
            ?.GetFirstChild<DocumentFormat.OpenXml.Drawing.EffectList>();
        if (fx == null) return null;
        var gfOrd = gfs.IndexOf(gf) + 1;
        return (fx.OuterXml, gfOrd);
    }

    // Connector <p:style> round-trip, mirroring GetShapeStyleXmlWithOrdinal.
    // A <p:cxnSp>'s theme-reference block (<a:lnRef>/<a:fillRef>/<a:effectRef>/
    // <a:fontRef>) has no typed Add/Set vocabulary and NodeBuilder doesn't
    // surface it, so the connector's styled line colour (commonly lnRef→accent1)
    // was silently dropped on dump∘replay — the line fell back to the theme's
    // default near-black stroke. Return the verbatim <p:style> XML + the
    // connector's ordinal within its parent scope so the emitter can raw-set
    // append it onto the replayed <p:cxnSp>.
    internal (string Xml, int CxnOrdinal)? GetConnectorStyleXmlWithOrdinal(string cxnPath)
    {
        var m = Regex.Match(cxnPath,
            @"^/slide\[(\d+)\]((?:/group\[(?:@id=)?\d+\])*)/connector\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        var byId = m.Groups[3].Value.Length > 0;
        var cxnIdx = int.Parse(m.Groups[4].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return null;
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(@id=)?(\d+)\]"))
        {
            var groupsHere = scope.Elements<GroupShape>().ToList();
            GroupShape? grp;
            if (gm.Groups[1].Value.Length > 0)
                grp = groupsHere.FirstOrDefault(g =>
                    g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                        == (uint)int.Parse(gm.Groups[2].Value));
            else
            {
                var gIdx = int.Parse(gm.Groups[2].Value);
                grp = (gIdx >= 1 && gIdx <= groupsHere.Count) ? groupsHere[gIdx - 1] : null;
            }
            if (grp == null) return null;
            scope = grp;
        }
        var cxns = scope.Elements<ConnectionShape>().ToList();
        ConnectionShape? cxn;
        int ordinal;
        if (byId)
        {
            cxn = cxns.FirstOrDefault(
                c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)cxnIdx);
            if (cxn == null) return null;
            ordinal = cxns.IndexOf(cxn) + 1;
        }
        else
        {
            if (cxnIdx < 1 || cxnIdx > cxns.Count) return null;
            cxn = cxns[cxnIdx - 1];
            ordinal = cxnIdx;
        }
        var styleEl = cxn.GetFirstChild<ShapeStyle>();
        if (styleEl == null) return null;
        return (styleEl.OuterXml, ordinal);
    }

    // Shape/placeholder image-fill TILE round-trip. ApplyShapeImageFill always
    // writes a stretched blipFill (<a:stretch>); a source shape whose image
    // fill tiles (<a:tile>, e.g. a body placeholder filled with a repeating
    // pattern) loses the tiling on replay and renders as one stretched image.
    // Return the source <a:tile> outer XML (plus the sp ordinal within its
    // parent scope) so EmitShape/EmitPlaceholder can raw-set replace the
    // replayed <a:stretch> with it. Null when the fill isn't a tiled blipFill.
    // Mirrors GetShapeStyleXmlWithOrdinal's path-resolution preamble verbatim.
    internal (string Xml, int SpOrdinal)? GetShapeBlipTileXmlWithOrdinal(string shapePath)
    {
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]((?:/group\[(?:@id=)?\d+\])*)/(?:shape|textbox|title|equation|placeholder)\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        var byId = m.Groups[3].Value.Length > 0;
        var shapeIdx = int.Parse(m.Groups[4].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return null;
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(@id=)?(\d+)\]"))
        {
            var groupsHere = scope.Elements<GroupShape>().ToList();
            GroupShape? grp;
            if (gm.Groups[1].Value.Length > 0)
                grp = groupsHere.FirstOrDefault(g =>
                    g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                        == (uint)int.Parse(gm.Groups[2].Value));
            else
            {
                var gIdx = int.Parse(gm.Groups[2].Value);
                grp = (gIdx >= 1 && gIdx <= groupsHere.Count) ? groupsHere[gIdx - 1] : null;
            }
            if (grp == null) return null;
            scope = grp;
        }
        var shapes = scope.Elements<Shape>().ToList();
        Shape? shape;
        int ordinal;
        if (byId)
        {
            shape = shapes.FirstOrDefault(
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)shapeIdx);
            if (shape == null) return null;
            ordinal = shapes.IndexOf(shape) + 1;
        }
        else
        {
            if (shapeIdx < 1 || shapeIdx > shapes.Count) return null;
            shape = shapes[shapeIdx - 1];
            ordinal = shapeIdx;
        }
        var blipFill = shape.ShapeProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.BlipFill>();
        if (blipFill == null) return null;
        var tile = blipFill.GetFirstChild<DocumentFormat.OpenXml.Drawing.Tile>();
        if (tile != null) return (tile.OuterXml, ordinal);
        // Bare blipFill: NEITHER <a:tile> nor <a:stretch>. PowerPoint renders
        // this mode-less form TILED (sample06), but ApplyShapeImageFill always
        // writes <a:stretch> on replay — the emitter must strip it back off.
        // Signal with an empty Xml so EmitBlipTileReplace emits a remove.
        if (blipFill.GetFirstChild<DocumentFormat.OpenXml.Drawing.Stretch>() == null)
            return ("", ordinal);
        return null;
    }

    // Text-field (<a:fld>) round-trip for a shape/placeholder. A slide-number /
    // date / footer placeholder renders its value from an <a:fld type="slidenum"
    // | "datetime…" | …> field run inside its text body. EmitParagraph only
    // re-emits <a:r>/<a:br> children, so the <a:fld> was dropped on replay —
    // the placeholder round-tripped but showed nothing (no page number). Return
    // the sp ordinal plus, per 1-based paragraph, the verbatim <a:fld> outer XML
    // so the emitter can raw-set insertbefore the paragraph's endParaRPr. Null
    // when the shape carries no fields. Mirrors GetShapeStyleXmlWithOrdinal.
    internal (int SpOrdinal, List<(int ParaOrdinal, string FldXml)> Fields)? GetShapeParagraphFieldXmls(string shapePath)
    {
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]((?:/group\[(?:@id=)?\d+\])*)/(?:shape|textbox|title|equation|placeholder)\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        var byId = m.Groups[3].Value.Length > 0;
        var shapeIdx = int.Parse(m.Groups[4].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return null;
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(@id=)?(\d+)\]"))
        {
            var groupsHere = scope.Elements<GroupShape>().ToList();
            GroupShape? grp;
            if (gm.Groups[1].Value.Length > 0)
                grp = groupsHere.FirstOrDefault(g =>
                    g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                        == (uint)int.Parse(gm.Groups[2].Value));
            else
            {
                var gIdx = int.Parse(gm.Groups[2].Value);
                grp = (gIdx >= 1 && gIdx <= groupsHere.Count) ? groupsHere[gIdx - 1] : null;
            }
            if (grp == null) return null;
            scope = grp;
        }
        var shapes = scope.Elements<Shape>().ToList();
        Shape? shape;
        int ordinal;
        if (byId)
        {
            shape = shapes.FirstOrDefault(
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)shapeIdx);
            if (shape == null) return null;
            ordinal = shapes.IndexOf(shape) + 1;
        }
        else
        {
            if (shapeIdx < 1 || shapeIdx > shapes.Count) return null;
            shape = shapes[shapeIdx - 1];
            ordinal = shapeIdx;
        }
        if (shape.TextBody == null) return null;
        var fields = new List<(int, string)>();
        int paraOrd = 0;
        foreach (var para in shape.TextBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
        {
            paraOrd++;
            foreach (var fld in para.Elements<DocumentFormat.OpenXml.Drawing.Field>())
                fields.Add((paraOrd, fld.OuterXml));
        }
        if (fields.Count == 0) return null;
        return (ordinal, fields);
    }

    // Resolve a shape path's blipFill image bytes (image fill on a non-Picture
    // <p:sp>). Mirrors GetImageBinary but walks <p:sp> (Shape) instead of
    // <p:pic> (Picture). Accepts /slide[N]/shape[K] and /slide[N]/shape[@id=ID]
    // path forms (group-nested shape paths are NOT supported in this initial
    // pass — covers the common slide-level blipFill case used by examples).
    public (byte[] Bytes, string ContentType)? GetShapeImageFillBinary(string shapePath)
    {
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]/shape\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var byId = m.Groups[2].Success;
        var idOrIdx = int.Parse(m.Groups[3].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return null;
        var shapes = shapeTree.Elements<Shape>().ToList();
        Shape? shape = null;
        if (byId)
        {
            shape = shapes.FirstOrDefault(s =>
            {
                var sid = s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                return sid.HasValue && sid.Value == (uint)idOrIdx;
            });
        }
        else
        {
            if (idOrIdx >= 1 && idOrIdx <= shapes.Count) shape = shapes[idOrIdx - 1];
        }
        if (shape == null) return null;
        var blipFill = shape.ShapeProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.BlipFill>();
        var embedId = blipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>()?.Embed?.Value;
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
