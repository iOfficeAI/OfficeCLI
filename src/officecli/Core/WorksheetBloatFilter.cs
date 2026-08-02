// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeCli.Core;

/// <summary>
/// Open-time mitigation for issue #149: worksheets that declare millions of
/// empty <c>&lt;c r="A5"/&gt;</c> cell elements (no value, no formula, no
/// style) blow up the Open XML SDK DOM — roughly 400 bytes of heap per
/// 17-byte XML token — turning an 11 MB file into multiple GB of RSS.
///
/// Common spreadsheet applications treat such cells as nonexistent:
/// they are dropped on load and never written back — a style-less empty
/// cell carries no data. Removing them is therefore semantically lossless.
///
/// A cell is removable ONLY when all three hold:
///   1. it has an <c>r</c> attribute,
///   2. <c>r</c> is its ONLY attribute (no s/t/cm/vm/ph, no xmlns decls),
///   3. it has no child nodes at all (no v/f/is, not even whitespace text).
///
/// Rule 1 matters: OOXML allows cells WITHOUT <c>r</c>, whose column is
/// implied by their position in the row. A bare <c>&lt;c/&gt;</c> is a
/// position placeholder — removing it would silently shift every following
/// r-less cell one column left. Same for rows: a <c>&lt;row&gt;</c> without
/// <c>r</c> is sequence-positioned and is never removed.
///
/// The filter is gated: only worksheet parts whose uncompressed size
/// exceeds <see cref="SheetSizeThreshold"/> are scanned, and the package
/// is only swapped when at least one cell was actually removed — normal
/// files keep the existing open path byte-for-byte.
/// </summary>
public static class WorksheetBloatFilter
{
    /// <summary>
    /// Uncompressed worksheet part size that triggers a filter scan.
    /// 4 MB ≈ 250k declared cells — far beyond any hand-authored sheet,
    /// cheap enough that one extra streaming parse is negligible next to
    /// the SDK DOM parse the file is about to pay anyway.
    /// </summary>
    public const long SheetSizeThreshold = 4 * 1024 * 1024;

    /// <summary>
    /// Dimension-rewrite guard: re-parsing the filtered output to patch
    /// &lt;dimension&gt; is only done when the output is small enough that
    /// an XDocument round-trip is trivially cheap.
    /// </summary>
    private const long DimensionRewriteMaxBytes = 64 * 1024 * 1024;

    public sealed class FilterResult
    {
        /// <summary>Filtered package as a fresh zip; null = no filtering applied.</summary>
        public MemoryStream? Package { get; init; }
        public long RemovedCells { get; init; }
        public IReadOnlyList<string> FilteredParts { get; init; } = Array.Empty<string>();

        public static readonly FilterResult None = new();
    }

    /// <summary>
    /// Scans the package on <paramref name="packageStream"/> (an open,
    /// seekable stream over the .xlsx zip). If any worksheet part exceeds
    /// the size threshold AND contains removable bare cells, returns a
    /// filtered copy of the whole package; otherwise returns
    /// <see cref="FilterResult.None"/>. The input stream is left at
    /// position 0 either way and is never written to.
    /// </summary>
    public static FilterResult TryFilter(Stream packageStream)
    {
        packageStream.Position = 0;
        try
        {
            using var zin = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true);

            if (!zin.Entries.Any(e => IsWorksheetEntry(e.FullName) && e.Length >= SheetSizeThreshold))
                return FilterResult.None;

            var outMs = new MemoryStream();
            long removedTotal = 0;
            var filteredParts = new List<string>();

            using (var zout = new ZipArchive(outMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var entry in zin.Entries)
                {
                    // OPC packages have no directory entries, but tolerate them.
                    if (entry.FullName.EndsWith('/')) continue;

                    var newEntry = zout.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                    using var os = newEntry.Open();
                    using var ins = entry.Open();

                    if (IsWorksheetEntry(entry.FullName) && entry.Length >= SheetSizeThreshold)
                    {
                        var removed = FilterSheetStream(ins, os);
                        if (removed > 0)
                        {
                            removedTotal += removed;
                            filteredParts.Add(entry.FullName);
                        }
                    }
                    else
                    {
                        ins.CopyTo(os);
                    }
                }
            }

            if (removedTotal == 0)
            {
                outMs.Dispose();
                return FilterResult.None;
            }

            outMs.Position = 0;
            return new FilterResult
            {
                Package = outMs,
                RemovedCells = removedTotal,
                FilteredParts = filteredParts,
            };
        }
        catch (InvalidDataException)
        {
            // Not a readable zip — let the SDK produce its own error on open.
            return FilterResult.None;
        }
        finally
        {
            if (packageStream.CanSeek) packageStream.Position = 0;
        }
    }

    private static bool IsWorksheetEntry(string fullName) =>
        fullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
        && fullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        && !fullName.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Streams one worksheet part from <paramref name="input"/> to
    /// <paramref name="output"/>, dropping removable bare cells (and rows
    /// that become empty and carry nothing but r/spans). Rows are buffered
    /// one at a time (O(row) memory), everything outside sheetData is
    /// copied verbatim. Returns the number of cells removed.
    /// </summary>
    public static long FilterSheetStream(Stream input, Stream output)
    {
        // Filter into a buffer first: <dimension> precedes <sheetData> in
        // the schema, so its corrected value is only known afterwards.
        using var buffer = new MemoryStream();
        var stats = FilterCore(input, buffer);

        if (stats.RemovedCells > 0
            && !stats.DimensionUnsafe
            && buffer.Length <= DimensionRewriteMaxBytes)
        {
            buffer.Position = 0;
            RewriteDimension(buffer, output, stats);
        }
        else
        {
            buffer.Position = 0;
            buffer.CopyTo(output);
        }
        return stats.RemovedCells;
    }

    private sealed class SheetStats
    {
        public long RemovedCells;
        // Bounds of surviving cells, for the dimension rewrite.
        public int MinRow = int.MaxValue, MaxRow, MinCol = int.MaxValue, MaxCol;
        // True when a surviving cell has no r attribute: its position is
        // sequence-implied and we cannot cheaply know the real bounds, so
        // the dimension element is left untouched (stale dimensions are
        // tolerated by common spreadsheet applications and this codebase alike).
        public bool DimensionUnsafe;

        public bool HasBounds => MinRow != int.MaxValue;
    }

    private static SheetStats FilterCore(Stream input, Stream buffer)
    {
        var stats = new SheetStats();
        var readerSettings = new XmlReaderSettings
        {
            IgnoreWhitespace = false,
            IgnoreComments = false,
            DtdProcessing = DtdProcessing.Prohibit,
        };
        using var reader = XmlReader.Create(input, readerSettings);
        var writerSettings = new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(false),
            CloseOutput = false,
        };
        using var writer = XmlWriter.Create(buffer, writerSettings);

        string? rootNs = null;
        var inSheetData = false;
        var sheetDataDepth = -1;
        // XNode.ReadFrom consumes the row element AND leaves the reader
        // positioned on the node that follows it — that node must be
        // processed without another Read(), or one node per row is lost
        // (eventually an end tag, desyncing the writer state machine).
        var skipRead = false;

        while (skipRead || reader.Read())
        {
            skipRead = false;
            switch (reader.NodeType)
            {
                case XmlNodeType.Element:
                    rootNs ??= reader.NamespaceURI; // worksheet root

                    if (inSheetData
                        && reader.Depth == sheetDataDepth + 1
                        && reader.LocalName == "row"
                        && reader.NamespaceURI == rootNs)
                    {
                        // Buffer one row, filter it, emit if it survives.
                        var rowEl = (XElement)XNode.ReadFrom(reader);
                        if (ProcessRow(rowEl, XNamespace.Get(rootNs), stats))
                            rowEl.WriteTo(writer);
                        skipRead = !reader.EOF;
                        continue;
                    }

                    if (!inSheetData
                        && reader.LocalName == "sheetData"
                        && reader.NamespaceURI == rootNs
                        && !reader.IsEmptyElement)
                    {
                        inSheetData = true;
                        sheetDataDepth = reader.Depth;
                    }

                    writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                    writer.WriteAttributes(reader, defattr: false);
                    if (reader.IsEmptyElement) writer.WriteEndElement();
                    break;

                case XmlNodeType.EndElement:
                    if (inSheetData && reader.Depth == sheetDataDepth)
                        inSheetData = false;
                    writer.WriteFullEndElement();
                    break;

                case XmlNodeType.Text:
                    writer.WriteString(reader.Value);
                    break;
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    writer.WriteWhitespace(reader.Value);
                    break;
                case XmlNodeType.CDATA:
                    writer.WriteCData(reader.Value);
                    break;
                case XmlNodeType.Comment:
                    writer.WriteComment(reader.Value);
                    break;
                case XmlNodeType.ProcessingInstruction:
                    writer.WriteProcessingInstruction(reader.Name, reader.Value);
                    break;
                case XmlNodeType.EntityReference:
                    writer.WriteEntityRef(reader.Name);
                    break;
                case XmlNodeType.XmlDeclaration:
                case XmlNodeType.DocumentType:
                    // Declaration is re-emitted by XmlWriter; DTDs are prohibited.
                    break;
            }
        }
        writer.Flush();
        return stats;
    }

    /// <summary>Returns true when the row should be kept.</summary>
    private static bool ProcessRow(XElement row, XNamespace ns, SheetStats stats)
    {
        var removedAny = false;
        foreach (var cell in row.Elements(ns + "c").ToList())
        {
            if (IsBareCell(cell))
            {
                cell.Remove();
                stats.RemovedCells++;
                removedAny = true;
            }
        }

        // Record surviving-cell bounds for the dimension rewrite.
        foreach (var cell in row.Elements(ns + "c"))
        {
            var r = cell.Attribute("r")?.Value;
            if (r == null || !TryParseCellRef(r, out var col, out var rowIdx))
            {
                stats.DimensionUnsafe = true;
                continue;
            }
            if (rowIdx < stats.MinRow) stats.MinRow = rowIdx;
            if (rowIdx > stats.MaxRow) stats.MaxRow = rowIdx;
            if (col < stats.MinCol) stats.MinCol = col;
            if (col > stats.MaxCol) stats.MaxCol = col;
        }

        if (row.Nodes().Any())
        {
            // Row survives. If we removed cells and it carries a spans hint,
            // recompute it from the remaining cells (Excel tolerates stale
            // spans, but a correct hint is one line of work here).
            if (removedAny)
                RecomputeSpans(row, ns);
            return true;
        }

        // Empty row: removable only when sequence-stable (has r) and carries
        // no meaningful attributes (height, hidden, style, outline, …).
        var hasR = false;
        foreach (var attr in row.Attributes())
        {
            if (attr.IsNamespaceDeclaration) return true; // conservative
            if (attr.Name.Namespace != XNamespace.None) return true; // e.g. x14ac:dyDescent
            switch (attr.Name.LocalName)
            {
                case "r": hasR = true; break;
                case "spans": break;
                default: return true;
            }
        }
        return !hasR;
    }

    private static bool IsBareCell(XElement cell)
    {
        if (cell.Nodes().Any()) return false; // any child node — even whitespace
        XAttribute? only = null;
        foreach (var attr in cell.Attributes())
        {
            if (only != null) return false; // more than one attribute
            only = attr;
        }
        return only != null
            && !only.IsNamespaceDeclaration
            && only.Name.Namespace == XNamespace.None
            && only.Name.LocalName == "r";
    }

    private static void RecomputeSpans(XElement row, XNamespace ns)
    {
        var spans = row.Attribute("spans");
        if (spans == null) return;
        int min = int.MaxValue, max = 0;
        foreach (var cell in row.Elements(ns + "c"))
        {
            var r = cell.Attribute("r")?.Value;
            if (r == null || !TryParseCellRef(r, out var col, out _))
                return; // r-less survivor — leave the hint alone
            if (col < min) min = col;
            if (col > max) max = col;
        }
        if (min != int.MaxValue)
            spans.Value = $"{min}:{max}";
    }

    private static void RewriteDimension(Stream filtered, Stream output, SheetStats stats)
    {
        var doc = XDocument.Load(filtered);
        var ns = doc.Root!.Name.Namespace;
        var dimension = doc.Root.Element(ns + "dimension");
        if (dimension != null)
        {
            dimension.SetAttributeValue("ref", stats.HasBounds
                ? $"{ColumnName(stats.MinCol)}{stats.MinRow}:{ColumnName(stats.MaxCol)}{stats.MaxRow}"
                : "A1");
        }
        var writerSettings = new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(false),
            CloseOutput = false,
        };
        using var writer = XmlWriter.Create(output, writerSettings);
        doc.Save(writer);
    }

    private static bool TryParseCellRef(string r, out int col, out int row)
    {
        col = 0; row = 0;
        var i = 0;
        while (i < r.Length && r[i] >= 'A' && r[i] <= 'Z')
        {
            col = col * 26 + (r[i] - 'A' + 1);
            i++;
        }
        if (i == 0 || i == r.Length) return false;
        for (; i < r.Length; i++)
        {
            if (r[i] < '0' || r[i] > '9') return false;
            row = row * 10 + (r[i] - '0');
        }
        return col > 0 && row > 0;
    }

    private static string ColumnName(int col)
    {
        var s = "";
        while (col > 0)
        {
            col--;
            s = (char)('A' + col % 26) + s;
            col /= 26;
        }
        return s;
    }
}
