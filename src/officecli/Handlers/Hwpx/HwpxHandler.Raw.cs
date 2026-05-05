// File: src/officecli/Handlers/Hwpx/HwpxHandler.Raw.cs
using System.IO.Compression;
using System.Xml.Linq;
using System.Xml.XPath;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Raw Layer ====================

    /// <summary>
    /// Return formatted XML string for the ZIP entry at partPath.
    /// partPath is a ZIP entry name e.g. "Contents/section0.xml", "Contents/header.xml".
    /// startRow/endRow/cols are ignored for HWPX (Excel compatibility params only).
    /// </summary>
    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        var entry = _doc.Archive.GetEntry(partPath)
            ?? throw new CliException($"Part not found: {partPath}") { Code = "not_found" };

        using var stream = entry.Open();
        var part = XDocument.Load(stream);
        return part.ToString();
    }

    /// <summary>
    /// Apply a mutation to the element selected by xpath within the ZIP entry at partPath.
    /// partPath = ZIP entry name (e.g. "Contents/section0.xml").
    /// xpath = XPath 1.0 expression relative to document root (e.g. "//*[local-name()='p'][1]").
    /// Actions: append | prepend | insertbefore | insertafter | replace | remove | setattr
    /// </summary>
    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        var entry = _doc.Archive.GetEntry(partPath)
            ?? throw new CliException($"Part not found: {partPath}") { Code = "not_found" };

        XDocument part;
        using (var readStream = entry.Open())
            part = XDocument.Load(readStream);

        // NOTE: XPath must use local-name() syntax for namespace-prefixed elements,
        // e.g. "//*[local-name()='p'][1]", because XPathSelectElement does not use
        // a namespace resolver by default.
        var element = part.XPathSelectElement(xpath)
            ?? throw new CliException($"XPath not found: {xpath}") { Code = "not_found" };

        switch (action.ToLowerInvariant())
        {
            case "remove":
                element.Remove();
                break;

            case "replace":
                if (string.IsNullOrEmpty(xml))
                    throw new CliException("replace action requires xml (XML fragment)")
                        { Code = "invalid_action" };
                element.ReplaceWith(XElement.Parse(xml));
                break;

            case "setattr":
                if (string.IsNullOrEmpty(xml))
                    throw new CliException("setattr action requires xml in format 'attrName=value'")
                        { Code = "invalid_action" };
                var eqIdx = xml.IndexOf('=');
                if (eqIdx <= 0)
                    throw new CliException("setattr xml must be in format 'attrName=value'")
                        { Code = "invalid_action" };
                element.SetAttributeValue(xml[..eqIdx], xml[(eqIdx + 1)..]);
                break;

            case "append":
                if (string.IsNullOrEmpty(xml))
                    throw new CliException("append action requires xml (XML fragment)")
                        { Code = "invalid_action" };
                element.Add(XElement.Parse(xml));
                break;

            case "prepend":
                if (string.IsNullOrEmpty(xml))
                    throw new CliException("prepend action requires xml (XML fragment)")
                        { Code = "invalid_action" };
                element.AddFirst(XElement.Parse(xml));
                break;

            case "insertbefore":
                if (string.IsNullOrEmpty(xml))
                    throw new CliException("insertbefore action requires xml (XML fragment)")
                        { Code = "invalid_action" };
                element.AddBeforeSelf(XElement.Parse(xml));
                break;

            case "insertafter":
                if (string.IsNullOrEmpty(xml))
                    throw new CliException("insertafter action requires xml (XML fragment)")
                        { Code = "invalid_action" };
                element.AddAfterSelf(XElement.Parse(xml));
                break;

            default:
                throw new CliException(
                    $"Unknown action: '{action}'. Valid actions: append, prepend, insertbefore, insertafter, replace, remove, setattr")
                { Code = "invalid_action" };
        }

        _dirty = true;
        // Write modified part back into ZIP archive (must be ZipArchiveMode.Update)
        // Delete-and-recreate pattern (avoids trailing bytes from SetLength(0))
        var writeEntry = _doc.Archive.GetEntry(partPath)
            ?? throw new CliException($"Part not found for write: {partPath}");
        var entryName = writeEntry.FullName;
        writeEntry.Delete();
        var newEntry = _doc.Archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writeStream = newEntry.Open();
        part.Save(writeStream);

        // Refresh in-memory DOM so subsequent get/query/set/validate see the raw edit
        RefreshCachedDocument(partPath, part);
    }

    /// <summary>
    /// After a raw ZIP write, synchronize the in-memory XDocument cache for the affected part.
    /// Prevents stale reads in resident mode where the same handler instance is reused.
    /// </summary>
    private void RefreshCachedDocument(string partPath, XDocument updatedDoc)
    {
        // Check if this is the header
        if (_doc.HeaderEntryPath != null
            && string.Equals(partPath, _doc.HeaderEntryPath, StringComparison.OrdinalIgnoreCase))
        {
            _doc.Header = updatedDoc;
            return;
        }

        // Check if this is a section
        var section = _doc.Sections.FirstOrDefault(s =>
            string.Equals(s.EntryPath, partPath, StringComparison.OrdinalIgnoreCase));
        if (section != null)
        {
            section.Document = updatedDoc;
        }
    }

    /// <summary>
    /// HWPX uses OPF packaging, NOT OPC. AddPart is meaningless for HWPX.
    /// Always throws CliException with unsupported_operation code.
    /// </summary>
    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType,
        Dictionary<string, string>? properties = null)
    {
        throw new CliException(
            "HWPX uses OPF packaging and does not support arbitrary part addition. " +
            "Use Raw() to modify existing XML entries directly.")
        {
            Code = "unsupported_operation",
            Suggestion = "Use 'raw' or 'raw-set' commands to modify existing HWPX XML content.",
            Help = "officecli raw document.hwpx Contents/section0.xml"
        };
    }
}
