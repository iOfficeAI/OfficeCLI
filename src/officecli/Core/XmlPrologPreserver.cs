using System.IO.Compression;
using System.Text;

namespace OfficeCli.Core;

/// <summary>
/// Preserves a part's XML prolog across a round-trip through the Open-XML SDK.
///
/// The SDK models a part as its root element. Comments and processing
/// instructions sitting *between* the XML declaration and the root element are
/// not part of that tree, so the first time the part is re-serialized — any
/// edit, or any editable save — they are gone. The declaration itself
/// survives, because the SDK re-emits it, which is what makes the loss easy to
/// miss: the saved part is well-formed, the command exits 0, and nothing on the
/// read surface (`raw` returns the root element) mentions what was dropped.
///
/// So the prolog is captured from the package as it was opened, and re-attached
/// once the SDK has written the new part. Both halves are best-effort, and both
/// touch an entry only when it actually carried prolog content — a package
/// whose parts have the usual declaration-then-root shape is left exactly as
/// the SDK wrote it.
/// </summary>
internal static class XmlPrologPreserver
{
    /// <summary>
    /// Bytes read from the head of a part when looking for its prolog. A prolog
    /// sits before the root element and is empty in practice; the cap only
    /// keeps a pathological part (a multi-megabyte single line) from being
    /// pulled into memory to find that out.
    /// </summary>
    private const int MaxHeadBytes = 16 * 1024;

    /// <summary>
    /// Collect the prolog of every XML part in <paramref name="package"/>.
    /// Returns null when no part has one — the caller then has nothing to
    /// restore, and the save path is left byte-for-byte alone. Leaves the
    /// stream position as it found it.
    /// </summary>
    public static Dictionary<string, string>? Capture(Stream package)
    {
        var found = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var pos = package.Position;
            package.Position = 0;
            using (var zip = new ZipArchive(package, ZipArchiveMode.Read, leaveOpen: true))
            {
                foreach (var entry in zip.Entries)
                {
                    if (!IsXmlPart(entry.FullName)) continue;
                    var head = ReadHead(entry);
                    if (head == null) continue;
                    if (!TryScanHead(head, out var prologStart, out var rootStart)) continue;
                    var prolog = head[prologStart..rootStart];
                    // Whitespace between the declaration and the root is not
                    // content — a consumer loses nothing when the SDK drops it,
                    // and treating it as prolog would force a rewrite of every
                    // part that was ever pretty-printed.
                    if (prolog.Trim().Length == 0) continue;
                    found[entry.FullName] = prolog;
                }
            }
            package.Position = pos;
        }
        catch
        {
            // An unreadable package simply gets no preservation; the caller's
            // own error handling owns the real diagnosis.
        }
        return found.Count == 0 ? null : found;
    }

    /// <summary>
    /// Re-attach each captured prolog to <paramref name="path"/> (the
    /// just-written package, still a temp file). Each entry is rewritten only
    /// when the write in front of us actually lost the prolog, so an entry the
    /// SDK copied through untouched keeps its own bytes.
    /// </summary>
    public static void Restore(string path, Dictionary<string, string>? prologs)
    {
        if (prologs == null || prologs.Count == 0) return;
        if (!File.Exists(path)) return;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite);
            using var zip = new ZipArchive(fs, ZipArchiveMode.Update, leaveOpen: false);
            foreach (var (entryName, prolog) in prologs)
            {
                var entry = zip.GetEntry(entryName);
                if (entry == null) continue;
                string xml;
                using (var rs = entry.Open())
                using (var sr = new StreamReader(rs))
                    xml = sr.ReadToEnd();
                if (!TryScanHead(xml, out var prologStart, out var rootStart)) continue;
                // The write already carries a prolog: the part was never
                // modelled (copied through verbatim), so re-inserting would
                // duplicate it. Nothing to do.
                if (xml[prologStart..rootStart].Trim().Length != 0) continue;

                var patched = xml[..prologStart] + prolog + xml[rootStart..];
                entry.Delete();
                var fresh = zip.CreateEntry(entryName);
                using var ws = fresh.Open();
                using var sw = new StreamWriter(ws, new UTF8Encoding(false));
                sw.Write(patched);
            }
        }
        catch
        {
            // Best-effort, matching the neighbour rewrites in the save path: a
            // locked or malformed package keeps the SDK's bytes rather than
            // failing the save the caller already reported as successful.
        }
    }

    private static bool IsXmlPart(string name)
        => name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

    private static string? ReadHead(ZipArchiveEntry entry)
    {
        try
        {
            using var s = entry.Open();
            var buf = new byte[(int)Math.Min(entry.Length, MaxHeadBytes)];
            var n = 0;
            while (n < buf.Length)
            {
                var r = s.Read(buf, n, buf.Length - n);
                if (r <= 0) break;
                n += r;
            }
            return Encoding.UTF8.GetString(buf, 0, n);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Scan a document head. On success <paramref name="prologStart"/> is the
    /// offset just past the BOM and the XML declaration (where a prolog begins,
    /// and where a captured one goes back), and <paramref name="rootStart"/> is
    /// the offset of the root element's opening '&lt;'. Returns false when the
    /// head is not the declaration / misc / root shape — the part is then left
    /// exactly as the SDK wrote it.
    /// </summary>
    private static bool TryScanHead(string xml, out int prologStart, out int rootStart)
    {
        prologStart = rootStart = 0;
        var i = 0;
        // BOM. The SDK writes one; a re-encoded inner segment may not.
        if (i < xml.Length && xml[i] == '\uFEFF') i++;

        // XML declaration. `<?xml` counts only when followed by whitespace or
        // `?>`, so the legal `<?xml-stylesheet?>` PI is not mistaken for it.
        if (StartsAt(xml, i, "<?xml")
            && i + 5 < xml.Length
            && (char.IsWhiteSpace(xml[i + 5])
                || (xml[i + 5] == '?' && i + 6 < xml.Length && xml[i + 6] == '>')))
        {
            var end = xml.IndexOf("?>", i + 5, StringComparison.Ordinal);
            if (end < 0) return false;
            i = end + 2;
        }
        prologStart = i;

        // Misc: whitespace, comments, processing instructions — everything a
        // prolog is allowed to hold before the root element. A DOCTYPE is
        // deliberately not traversed: DTDs are prohibited in OOXML (the same
        // reason WorksheetBloatFilter drops them), and its internal subset
        // needs real parsing to skip safely. The scan stops there and leaves the
        // part as the SDK wrote it.
        while (i < xml.Length)
        {
            if (char.IsWhiteSpace(xml[i])) { i++; continue; }
            if (StartsAt(xml, i, "<!--"))
            {
                var end = xml.IndexOf("-->", i + 4, StringComparison.Ordinal);
                if (end < 0) return false;
                i = end + 3;
                continue;
            }
            if (StartsAt(xml, i, "<?"))
            {
                var end = xml.IndexOf("?>", i + 2, StringComparison.Ordinal);
                if (end < 0) return false;
                i = end + 2;
                continue;
            }
            break;
        }

        // What is left must be the root element's opening tag.
        if (i + 1 >= xml.Length || xml[i] != '<' || xml[i + 1] == '/' || xml[i + 1] == '!')
            return false;
        rootStart = i;
        return true;
    }

    private static bool StartsAt(string s, int i, string token)
        => i + token.Length <= s.Length && string.CompareOrdinal(s, i, token, 0, token.Length) == 0;
}
