// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler : IDocumentHandler, Rendering.IRenderModelHost
{
    private readonly SpreadsheetDocument _doc;

    object? Rendering.IRenderModelHost.RenderModel => _doc;
    private readonly string _filePath;
    private readonly HashSet<string> _initialSheetNames;
    private readonly HashSet<WorksheetPart> _dirtyWorksheets = new();
    private bool _dirtyStylesheet;
    private bool _disposed;
    // Backing FileStream — mirrors the PPT/Word pattern. Opening via a shared
    // FileStream (FileShare.Read in editable mode) lets external readers
    // observe the file while the handler is alive, which is required for
    // mid-session `save` snapshots to be useful to third-party consumers
    // (issue #114).
    private FileStream? _backingStream;
    // Issue #149: when the package contained bloat-filtered worksheets the
    // SDK is opened over this slimmed in-memory copy instead of the backing
    // FileStream (which stays open purely for locking + write-back).
    // See Core/WorksheetBloatFilter.cs for the removal rules and why this
    // is semantically lossless.
    private MemoryStream? _filteredPackageStream;
    private readonly bool _editable;
    // Row index cache: SheetData → sorted map of rowIndex → Row.
    // Turns the O(n) linear scan in FindOrCreateCell into O(1) lookup + O(log n) insert.
    // Invalidated by InvalidateRowIndex() whenever rows are structurally modified (shift, remove).
    private Dictionary<SheetData, SortedList<uint, Row>>? _rowIndex;
    public int LastFindMatchCount { get; internal set; }
    // Number of elements a no-slash selector Set matched and mutated (Sheet1!row[...]).
    // Read by the CLI/resident to echo the multi-element change count.
    public int LastSelectorSetCount { get; internal set; }

    /// <summary>
    /// Set true by Add/Set/Remove/RawSet, consumed by Save/Dispose to decide
    /// whether to stamp <c>docProps/custom.xml</c> with an OfficeCLI audit
    /// trail. Pure Get/Query sessions leave this false.
    /// </summary>
    internal bool Modified { get; set; }

    /// <summary>
    /// Number of bare empty cell declarations removed at open by
    /// <see cref="OfficeCli.Core.WorksheetBloatFilter"/> (issue #149).
    /// Zero for normal files.
    /// </summary>
    public long BloatCellsRemoved { get; private set; }

    public ExcelHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        _editable = editable;
        try
        {
            var share = editable ? FileShare.Read : FileShare.ReadWrite;
            var access = editable ? FileAccess.ReadWrite : FileAccess.Read;
            _backingStream = new FileStream(filePath, FileMode.Open, access, share);

            // Issue #149: sheets that declare millions of empty cells make
            // the SDK DOM balloon to GBs. Filter them out (lossless —
            // spreadsheet applications discard such cells on load) before the
            // SDK ever parses the part. Gated: normal files take the
            // original path untouched.
            var bloat = OfficeCli.Core.WorksheetBloatFilter.TryFilter(_backingStream);
            if (bloat.Package != null)
            {
                _filteredPackageStream = bloat.Package;
                BloatCellsRemoved = bloat.RemovedCells;
                try
                {
                    Console.Error.WriteLine(
                        $"warning: {Path.GetFileName(filePath)} declares {bloat.RemovedCells:N0} empty cells " +
                        $"with no value, formula or style ({string.Join(", ", bloat.FilteredParts)}); " +
                        "they were skipped on load and will be omitted on save (other spreadsheet applications do the same).");
                }
                catch { /* stderr unavailable (resident) — property still set */ }
            }
            else if (editable)
            {
                // Crash-atomic saves: an editable session works against an
                // in-memory copy of the package so the on-disk file is only ever
                // swapped atomically (temp + File.Replace in
                // WriteBackFilteredPackage), never rewritten in place. Without
                // this, _doc.Save() rewrites the backing file directly, and a
                // process death mid-write leaves a truncated, unopenable file with
                // no fallback (the original bytes are already gone). Read-only
                // sessions skip this and keep streaming from the FileStream — they
                // never write, so there is nothing to make atomic and no reason to
                // pay the in-memory copy. The bloat-filter branch above already
                // routes editable bloated files through the same in-memory model.
                _backingStream.Position = 0;
                var mem = new MemoryStream();
                _backingStream.CopyTo(mem);
                mem.Position = 0;
                _filteredPackageStream = mem;
            }

            _doc = SpreadsheetDocument.Open(
                (Stream?)_filteredPackageStream ?? _backingStream, editable);
            // Force early validation: access WorkbookPart to catch corrupt packages now
            _ = _doc.WorkbookPart?.Workbook;
            // Capture initial sheet names to detect duplicate additions
            _initialSheetNames = new HashSet<string>(
                GetWorkbook().GetFirstChild<Sheets>()?.Elements<Sheet>()
                    .Select(s => s.Name?.Value ?? "")
                    .Where(n => !string.IsNullOrEmpty(n)) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
        }
        catch (DocumentFormat.OpenXml.Packaging.OpenXmlPackageException ex)
        {
            throw new InvalidOperationException(
                $"Cannot open {Path.GetFileName(filePath)}: {ex.Message}", ex);
        }
    }

    // ==================== Raw Layer ====================

    // CONSISTENCY(zip-uri-lookup): any partPath ending in `.xml` is treated
    // as a literal zip-internal URI and resolved via `RawXmlHelper.FindPartByZipUri`.
    // This replaces the old hand-curated alias table (workbook/styles/...) which
    // could never cover everything — sheet/slide-scoped parts, footnotes,
    // custom XML, etc. were all unreachable. Semantic short names (`/workbook`,
    // `/Sheet1`, `/chart[N]`) continue to route through the switch below.

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var workbookPart = _doc.WorkbookPart;
        if (workbookPart == null) return "(empty)";

        // Zip-URI form: any path ending in .xml or .rels is resolved literally
        // against the package. No alias table needed.
        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            // CONSISTENCY(zip-uri-row-filter): if the resolved part is a
            // worksheet AND the caller asked for row/column filtering,
            // route through the same filter as the semantic /SheetName
            // path. Without this, --start/--end/--cols would be silently
            // ignored on zip-URI worksheet reads.
            if ((startRow.HasValue || endRow.HasValue || cols != null)
                && RawXmlHelper.FindPartByZipUri(_doc, partPath) is WorksheetPart wsp)
            {
                return RawSheetWithFilter(wsp, startRow, endRow, cols);
            }

            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/workbook, /Sheet1, /chart[N]) for stable identification.");
            return xml;
        }

        if (partPath == "/" || partPath == "/workbook")
            return workbookPart.Workbook?.OuterXml ?? "(empty)";

        if (partPath == "/styles")
        {
            // Raw is read-only; do not create the part if missing (would fail
            // when the package is opened read-only).
            var stylesPart = workbookPart.WorkbookStylesPart;
            return stylesPart?.Stylesheet?.OuterXml ?? "(no styles)";
        }

        if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            return sst?.SharedStringTable?.OuterXml ?? "(no shared strings)";
        }

        if (partPath == "/theme")
        {
            var themePart = workbookPart.ThemePart;
            return themePart?.Theme?.OuterXml ?? "(no theme)";
        }

        // Drawing part: /SheetName/drawing
        var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
        if (drawingMatch.Success)
        {
            var drawSheetName = drawingMatch.Groups[1].Value;
            var drawWs = FindWorksheet(drawSheetName)
                ?? throw SheetNotFoundException(drawSheetName);
            var dp = drawWs.DrawingsPart
                ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
            return dp.WorksheetDrawing!.OuterXml;
        }

        // Chart part: /SheetName/chart[N] or /chart[N]
        var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
        if (chartMatch.Success)
        {
            var chartSheetName = chartMatch.Groups[1].Value;
            var chartIdx = int.Parse(chartMatch.Groups[2].Value);
            var chartWs = FindWorksheet(chartSheetName)
                ?? throw SheetNotFoundException(chartSheetName);
            // Resolve via GetExcelCharts (document order, both legacy ChartPart and
            // extended cx ChartPart) so `raw` reaches the same charts `query`/`get`
            // list. The legacy-only GetChartPart missed funnel/treemap/sunburst/
            // boxWhisker/waterfall/histogram (all ExtendedChartPart) → "out of range".
            return GetChartSpaceOuterXml(chartWs.DrawingsPart, chartIdx);
        }

        // Global chart: /chart[N] — searches all sheets
        var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
        if (globalChartMatch.Success)
        {
            var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
            var all = new List<ExcelChartInfo>();
            foreach (var (_, wsp) in GetWorksheets())
                if (wsp.DrawingsPart != null) all.AddRange(ChartsForRaw(wsp.DrawingsPart));
            if (all.Count == 0)
                throw new ArgumentException("No charts found in workbook");
            return ChartSpaceOuterXmlAt(all, chartIdx);
        }

        // Try as sheet name
        var sheetName = partPath.TrimStart('/');
        var worksheet = FindWorksheet(sheetName);
        if (worksheet != null)
        {
            if (startRow.HasValue || endRow.HasValue || cols != null)
                return RawSheetWithFilter(worksheet, startRow, endRow, cols);
            return GetSheet(worksheet).OuterXml;
        }

        // /SheetName/<relId> fallback — resolve a worksheet relationship by id
        // (covers OLE embed parts, image parts, etc. that have no named path).
        // Open XML SDK generates relIds like "rId12" or "Rff3244f593f8481a";
        // accept both forms (any non-slash token starting with R/r).
        var relIdMatch = Regex.Match(partPath, @"^/([^/]+)/([Rr][A-Za-z0-9]+)$");
        if (relIdMatch.Success)
        {
            var relSheetName = relIdMatch.Groups[1].Value;
            var relId = relIdMatch.Groups[2].Value;
            var relWs = FindWorksheet(relSheetName)
                ?? throw SheetNotFoundException(relSheetName);
            try
            {
                var part = relWs.GetPartById(relId);
                if (part != null)
                {
                    var ct = part.ContentType ?? "";
                    bool isText = ct.Contains("xml", StringComparison.OrdinalIgnoreCase)
                        || ct.StartsWith("text/", StringComparison.OrdinalIgnoreCase);
                    using var partStream = part.GetStream();
                    if (isText)
                    {
                        using var reader = new StreamReader(partStream);
                        return reader.ReadToEnd();
                    }
                    long size = 0;
                    try { size = partStream.Length; } catch { /* non-seekable */ }
                    return $"(binary part: {ct}, {size} bytes)";
                }
            }
            catch (KeyNotFoundException)
            {
                // fall through to the unknown-part error
            }
            catch (ArgumentException)
            {
                // fall through to the unknown-part error
            }
        }

        throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /theme, /<SheetName>, /<SheetName>/drawing, /<SheetName>/chart[N], /chart[N], /<SheetName>/<relId>");
    }

    private static string RawSheetWithFilter(WorksheetPart worksheetPart, int? startRow, int? endRow, HashSet<string>? cols)
    {
        var worksheet = GetSheet(worksheetPart);
        var sheetData = worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
            return worksheet.OuterXml;

        var cloned = (Worksheet)worksheet.CloneNode(true);
        var clonedSheetData = cloned.GetFirstChild<SheetData>()!;
        clonedSheetData.RemoveAllChildren();

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowNum = (int)row.RowIndex!.Value;
            if (startRow.HasValue && rowNum < startRow.Value) continue;
            if (endRow.HasValue && rowNum > endRow.Value) break;

            if (cols != null)
            {
                var filteredRow = (Row)row.CloneNode(false);
                filteredRow.RowIndex = row.RowIndex;
                foreach (var cell in row.Elements<Cell>())
                {
                    var colName = ParseCellReference(cell.CellReference?.Value ?? "A1").Column;
                    if (cols.Contains(colName))
                        filteredRow.AppendChild(cell.CloneNode(true));
                }
                clonedSheetData.AppendChild(filteredRow);
            }
            else
            {
                clonedSheetData.AppendChild(row.CloneNode(true));
            }
        }

        return cloned.OuterXml;
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        Modified = true;
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        // Zip-URI form: resolve via package part tree, mutate the part's XML
        // stream directly (no SDK typed root needed — handles arbitrary XML
        // parts like footnotes, customXml, untyped sheet1.xml, etc.).
        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/workbook, /Sheet1, /chart[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;
        if (partPath is "/" or "/workbook")
        {
            rootElement = workbookPart.Workbook
                ?? throw new InvalidOperationException("No workbook");
        }
        else if (partPath == "/styles")
        {
            var styleManager = new ExcelStyleManager(workbookPart);
            rootElement = styleManager.EnsureStylesPart().Stylesheet!;
        }
        else if (partPath == "/sharedstrings")
        {
            var sst = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                ?? throw new InvalidOperationException("No shared strings");
            rootElement = sst.SharedStringTable!;
        }
        else if (partPath == "/theme")
        {
            rootElement = workbookPart.ThemePart?.Theme
                ?? throw new ArgumentException("No theme part");
        }
        else
        {
            // Drawing part: /SheetName/drawing
            var drawingMatch = Regex.Match(partPath, @"^/(.+)/drawing$");
            if (drawingMatch.Success)
            {
                var drawSheetName = drawingMatch.Groups[1].Value;
                var drawWs = FindWorksheet(drawSheetName)
                    ?? throw SheetNotFoundException(drawSheetName);
                var dp = drawWs.DrawingsPart
                    ?? throw new ArgumentException($"Sheet '{drawSheetName}' has no drawings");
                rootElement = dp.WorksheetDrawing!;
            }
            else
            {
            // Chart part: /SheetName/chart[N] or /chart[N]
            var chartMatch = Regex.Match(partPath, @"^/(.+)/chart\[(\d+)\]$");
            if (chartMatch.Success)
            {
                var chartSheetName = chartMatch.Groups[1].Value;
                var chartIdx = int.Parse(chartMatch.Groups[2].Value);
                var chartWs = FindWorksheet(chartSheetName)
                    ?? throw SheetNotFoundException(chartSheetName);
                // Resolve via GetExcelCharts so raw-set writes reach extended (cx)
                // charts too — mirrors the Raw() read path. GetChartPart was
                // legacy-only and 422'd cx charts with "out of range".
                rootElement = GetChartSpaceElement(chartWs.DrawingsPart, chartIdx);
            }
            else
            {
                var globalChartMatch = Regex.Match(partPath, @"^/chart\[(\d+)\]$");
                if (globalChartMatch.Success)
                {
                    var chartIdx = int.Parse(globalChartMatch.Groups[1].Value);
                    var all = new List<ExcelChartInfo>();
                    foreach (var (_, wsp) in GetWorksheets())
                        if (wsp.DrawingsPart != null) all.AddRange(ChartsForRaw(wsp.DrawingsPart));
                    if (all.Count == 0)
                        throw new ArgumentException("No charts found in workbook");
                    rootElement = ChartSpaceElementAt(all, chartIdx);
                }
                else
                {
                    // Try as sheet name
                    var sheetName = partPath.TrimStart('/');
                    var worksheet = FindWorksheet(sheetName)
                        ?? throw new ArgumentException($"Unknown part: {partPath}. Available: /workbook, /styles, /sharedstrings, /theme, /<SheetName>, /<SheetName>/chart[N], /chart[N]");
                    rootElement = GetSheet(worksheet);
                }
            }
            }
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // BUG-R5-01: silent — CLI wrappers print their own structured message.
        _ = affected;
    }

    public List<ValidationError> Validate()
    {
        // Mutations defer worksheet child reordering (CF / mergeCells /
        // hyperlinks / rowBreaks land in insertion order) to FlushDirtyParts,
        // which Save() runs before writing. Validating the raw in-memory tree
        // mid-session therefore reported schema-sequence errors on documents
        // that were perfectly fine once saved. Flush first, same as Save().
        FlushDirtyParts();
        return RawXmlHelper.ValidateDocument(_doc, _filePath);
    }

    public void Save()
    {
        // Excel mutations defer worksheet/stylesheet writes to FlushDirtyParts
        // (see ExcelHandler.Helpers.cs). Flush them first so the package
        // reflects the in-memory tree, then push the package out to the
        // backing stream and force the OS buffer to disk.
        FlushDirtyParts();
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        _doc.Save();
        WriteBackFilteredPackage();
        _backingStream?.Flush();
    }

    // Issue #149: when the SDK operates on the bloat-filtered in-memory
    // copy, saves land in that MemoryStream — push the bytes out to the
    // backing FileStream so mid-session `save` snapshots (issue #114) and
    // final closes hit the disk file exactly like the direct-stream path.
    // Read-only sessions never write back: the disk file stays untouched.
    private void WriteBackFilteredPackage()
    {
        if (_filteredPackageStream == null || _backingStream == null || !_editable)
            return;
        // Crash-atomic write-back (temp + File.Replace). The old path truncated
        // the original in place (SetLength(0)) then copied the new bytes over it,
        // so a process death between the two left the original already gone and
        // only a partial file on disk — unrecoverable. See AtomicPackageWriter.
        OfficeCli.Core.AtomicPackageWriter.Flush(
            _filteredPackageStream, _filePath,
            releaseLock: () => { _backingStream!.Dispose(); _backingStream = null; },
            reopenLock: () => { _backingStream = new FileStream(_filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read); });
    }

    /// <summary>See <see cref="OfficeCli.Handlers.WordHandler.DiscardOnDispose"/> —
    /// atomic-batch rollback: drop the in-memory DOM without serializing.</summary>
    public bool DiscardOnDispose { get; set; }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (DiscardOnDispose)
        {
            // Never serialize the poisoned DOM. Close the disk-backed streams
            // first so the package's dispose-time autosave has nowhere to
            // write (the filtered-package path autosaves into its
            // MemoryStream, which is discarded — WriteBackFilteredPackage is
            // the only disk write there and is skipped).
            _backingStream?.Dispose();
            _backingStream = null;
            try { _doc.Dispose(); } catch { /* autosave hit the closed stream — intended */ }
            _filteredPackageStream?.Dispose();
            _filteredPackageStream = null;
            return;
        }
        if (!Modified)
        {
            // A read-only inspection must be byte-preserving even when a caller
            // accidentally opened the handler in editable mode. Open XML SDK
            // serializes package relationships during Save/Dispose, which can
            // discard DrawingML hyperlink relationships it cannot round-trip.
            // Close the backing stream first so any dispose-time autosave has
            // nowhere to write, then discard the in-memory package.
            _backingStream?.Dispose();
            _backingStream = null;
            try { _doc.Dispose(); } catch { /* autosave hit the closed stream — intended */ }
            _filteredPackageStream?.Dispose();
            _filteredPackageStream = null;
            return;
        }
        try { FlushDirtyParts(); } catch { /* best-effort */ }
        // Mirror the PPT/Word pattern: when we own the backing FileStream the
        // package would otherwise leave the on-disk file in whatever state
        // the last auto-flush left it. Save explicitly before disposing.
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        try { _doc.Save(); } catch { /* read-only or already disposed */ }
        // Write back while the MemoryStream is guaranteed alive — the SDK
        // may close the stream it was opened over during _doc.Dispose().
        // After _doc.Save() the in-memory zip is complete (same contract
        // the direct FileStream path relies on for mid-session saves).
        try { WriteBackFilteredPackage(); } catch { /* best-effort */ }
        _doc.Dispose();
        _filteredPackageStream?.Dispose();
        _filteredPackageStream = null;
        _backingStream?.Dispose();
        _backingStream = null;
    }

}
