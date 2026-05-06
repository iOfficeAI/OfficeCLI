// File: src/officecli/Handlers/Hwpx/HwpxHandler.Validate.cs
using System.IO.Compression;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Validation ====================

    public List<ValidationError> Validate()
    {
        var errors = new List<ValidationError>();

        // Level 1: ZIP integrity
        if (!ValidateZipIntegrity(errors))
            return errors; // Critical — stop here

        // Level 2: OPF manifest
        if (!ValidateOpfManifest(errors))
            return errors; // Critical — stop here

        // Level 3: XML well-formedness
        ValidateXmlWellFormedness(errors);

        // Level 4: IDRef consistency (charPrIDRef on runs → header.xml charPr)
        ValidateIdRefConsistency(errors);

        // Level 5: Table structure
        ValidateTableStructure(errors);

        // Level 6: Namespace declarations
        ValidateNamespaceDeclarations(errors);

        // Level 7: BinData integrity (Plan 94 — merged from ViewAsIssues)
        ValidateBinDataIntegrity(errors);

        // Level 8: Field pair consistency (Plan 94)
        ValidateFieldPairs(errors);

        // Level 9: Section count consistency (Plan 94)
        ValidateSectionCount(errors);

        return errors;
    }

    /// <summary>
    /// Level 1: Verify the file is a valid ZIP archive.
    /// Returns false if ZIP is corrupted (critical failure).
    /// </summary>
    private bool ValidateZipIntegrity(List<ValidationError> errors)
    {
        try
        {
            // The archive is already open (loaded in constructor).
            // Verify we can enumerate entries without error.
            var entryCount = _doc.Archive.Entries.Count;
            if (entryCount == 0)
            {
                errors.Add(new ValidationError(
                    "zip_empty",
                    "ZIP archive contains no entries",
                    "/",
                    null));
                return false;
            }
            return true;
        }
        catch (InvalidDataException ex)
        {
            errors.Add(new ValidationError(
                "zip_corrupt",
                $"File is not a valid ZIP archive: {ex.Message}",
                "/",
                null));
            return false;
        }
    }

    /// <summary>
    /// Level 2: Verify OPF manifest structure.
    /// - mimetype entry must exist and be the first ZIP entry with no compression
    /// - META-INF/container.xml must be parseable
    /// </summary>
    private bool ValidateOpfManifest(List<ValidationError> errors)
    {
        bool critical = false;

        // Check mimetype entry
        var mimetypeEntry = _doc.Archive.GetEntry("mimetype");
        if (mimetypeEntry == null)
        {
            errors.Add(new ValidationError(
                "opf_missing_mimetype",
                "HWPX package missing 'mimetype' entry (required by OPF spec)",
                "/mimetype",
                null));
        }
        else
        {
            // mimetype must be first entry
            var firstEntry = _doc.Archive.Entries.FirstOrDefault();
            if (firstEntry?.FullName != "mimetype")
            {
                errors.Add(new ValidationError(
                    "opf_mimetype_not_first",
                    "mimetype must be the first ZIP entry (found: " + (firstEntry?.FullName ?? "none") + ")",
                    "/mimetype",
                    null));
            }

            // Plan 81: mimetype must use ZIP_STORED (no compression)
            // .NET ZipArchiveEntry has no CompressionMethod — use heuristic: CompressedLength == Length
            if (mimetypeEntry.CompressedLength != mimetypeEntry.Length)
            {
                errors.Add(new ValidationError(
                    "package_mimetype_compressed",
                    $"mimetype entry should use ZIP_STORED (no compression), but appears compressed (size={mimetypeEntry.Length}, compressed={mimetypeEntry.CompressedLength})",
                    "/mimetype",
                    null));
            }

            // G5: Validate mimetype content value
            using var mimeStream = mimetypeEntry.Open();
            using var mimeReader = new StreamReader(mimeStream, System.Text.Encoding.ASCII);
            var mimeContent = mimeReader.ReadToEnd().Trim();
            var validMimeTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "application/hwp+zip",
                "application/vnd.hancom.hwp",
                "application/vnd.hancom.hwpx",
                "application/haansofthwp"
            };
            if (!validMimeTypes.Contains(mimeContent))
            {
                errors.Add(new ValidationError(
                    "package_mimetype_invalid",
                    $"Unexpected MIME type '{mimeContent}' (expected one of: {string.Join(", ", validMimeTypes)})",
                    "/mimetype",
                    mimeContent));
            }
        }

        // Check META-INF/container.xml
        var containerEntry = _doc.Archive.GetEntry("META-INF/container.xml");
        if (containerEntry != null)
        {
            try
            {
                using var stream = containerEntry.Open();
                XDocument.Load(stream); // parse test
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "opf_container_invalid",
                    $"META-INF/container.xml is not valid XML: {ex.Message}",
                    "/META-INF/container.xml",
                    "container.xml"));
                critical = true;
            }
        }

        // Plan 81: Rootfile resolution check
        if (containerEntry != null && _doc.RootfilePath != null)
        {
            var rootEntry = _doc.Archive.GetEntry(_doc.RootfilePath);
            if (rootEntry == null)
            {
                errors.Add(new ValidationError(
                    "package_rootfile_missing",
                    $"container.xml rootfile points to '{_doc.RootfilePath}' but entry not found in archive",
                    "/META-INF/container.xml",
                    null));
                critical = true;
            }
        }

        // Plan 81: version.xml check (warning only). Real Hancom exports have
        // appeared with version.xml at the package root, so accept both forms.
        var versionEntry = _doc.Archive.GetEntry("Contents/version.xml")
            ?? _doc.Archive.GetEntry("version.xml");
        if (versionEntry == null)
        {
            errors.Add(new ValidationError(
                "package_version_missing",
                "Contents/version.xml not found (optional but recommended by OWPML spec)",
                "/Contents/version.xml",
                null));
            // Warning severity — not critical
        }

        // Plan 81: Section count consistency (manifest vs loaded)
        if (_doc.ManifestDoc != null)
        {
            var manifestSectionCount = _doc.ManifestDoc.Descendants()
                .Count(e => e.Attribute("media-type")?.Value?.Contains("section") ?? false);
            if (manifestSectionCount != _doc.Sections.Count && manifestSectionCount > 0)
            {
                errors.Add(new ValidationError(
                    "package_section_mismatch",
                    $"Manifest declares {manifestSectionCount} sections but {_doc.Sections.Count} loaded",
                    "/Contents/content.hpf",
                    null));
            }
        }

        // Check content.hpf (the OPF package file)
        var hpfEntry = _doc.Archive.GetEntry("Contents/content.hpf");
        if (hpfEntry == null)
        {
            errors.Add(new ValidationError(
                "opf_missing_hpf",
                "HWPX package missing 'Contents/content.hpf' manifest",
                "/Contents/content.hpf",
                null));
            critical = true;
        }
        else
        {
            try
            {
                using var stream = hpfEntry.Open();
                XDocument.Load(stream); // parse test
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "opf_hpf_invalid",
                    $"Contents/content.hpf is not valid XML: {ex.Message}",
                    "/Contents/content.hpf",
                    "content.hpf"));
                critical = true;
            }
        }

        return !critical;
    }

    /// <summary>
    /// Level 3: Verify all .xml entries in the archive parse without exception.
    /// </summary>
    private void ValidateXmlWellFormedness(List<ValidationError> errors)
    {
        foreach (var entry in _doc.Archive.Entries)
        {
            if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                continue;

            try
            {
                using var stream = entry.Open();
                XDocument.Load(stream);
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "xml_malformed",
                    $"XML parse error in '{entry.FullName}': {ex.Message}",
                    $"/{entry.FullName}",
                    entry.FullName));
            }
        }
    }

    /// <summary>
    /// Level 4: Verify all charPrIDRef values on hp:run elements reference
    /// existing charPr entries in header.xml.
    /// Scans ALL sections (not just PrimarySection).
    /// </summary>
    private void ValidateIdRefConsistency(List<ValidationError> errors)
    {
        if (_doc.Header?.Root == null)
        {
            // No header.xml — can't validate refs, but this is already flagged by ViewAsIssues
            return;
        }

        // Collect all valid charPr IDs from header.xml
        var validCharPrIds = new HashSet<string>(
            _doc.Header.Root
                .Descendants(HwpxNs.Hh + "charPr")
                .Select(cp => cp.Attribute("id")?.Value)
                .Where(id => id != null)!);

        // Scan ALL sections for charPrIDRef references
        foreach (var section in _doc.Sections)
        {
            int localParaIdx = 0;
            foreach (var para in section.Paragraphs)
            {
                localParaIdx++;
                int runIdx = 0;
                foreach (var run in para.Elements(HwpxNs.Hp + "run"))
                {
                    runIdx++;
                    var charPrIdRef = run.Attribute("charPrIDRef")?.Value;
                    if (charPrIdRef == null) continue;

                    if (!validCharPrIds.Contains(charPrIdRef))
                    {
                        var path = $"/section[{section.Index + 1}]/p[{localParaIdx}]/run[{runIdx}]";
                        errors.Add(new ValidationError(
                            "idref_dangling",
                            $"charPrIDRef=\"{charPrIdRef}\" references non-existent charPr in header.xml",
                            path,
                            section.EntryPath));
                    }
                }
            }
        }

        // Also validate paraPrIDRef references
        var validParaPrIds = new HashSet<string>(
            _doc.Header.Root
                .Descendants(HwpxNs.Hh + "paraPr")
                .Select(pp => pp.Attribute("id")?.Value)
                .Where(id => id != null)!);

        foreach (var section in _doc.Sections)
        {
            int localParaIdx = 0;
            foreach (var para in section.Paragraphs)
            {
                localParaIdx++;
                var paraPrIdRef = para.Attribute("paraPrIDRef")?.Value;
                if (paraPrIdRef == null) continue;

                if (!validParaPrIds.Contains(paraPrIdRef))
                {
                    var path = $"/section[{section.Index + 1}]/p[{localParaIdx}]";
                    errors.Add(new ValidationError(
                        "idref_dangling",
                        $"paraPrIDRef=\"{paraPrIdRef}\" references non-existent paraPr in header.xml",
                        path,
                        section.EntryPath));
                }
            }
        }

        // Validate styleIDRef references
        var validStyleIds = new HashSet<string>(
            _doc.Header.Root
                .Descendants(HwpxNs.Hh + "style")
                .Select(s => s.Attribute("id")?.Value)
                .Where(id => id != null)!);

        foreach (var section in _doc.Sections)
        {
            int localParaIdx = 0;
            foreach (var para in section.Paragraphs)
            {
                localParaIdx++;
                var styleIdRef = para.Attribute("styleIDRef")?.Value;
                if (styleIdRef == null) continue;

                if (!validStyleIds.Contains(styleIdRef))
                {
                    var path = $"/section[{section.Index + 1}]/p[{localParaIdx}]";
                    errors.Add(new ValidationError(
                        "idref_dangling",
                        $"styleIDRef=\"{styleIdRef}\" references non-existent style in header.xml",
                        path,
                        section.EntryPath));
                }
            }
        }
    }

    /// <summary>
    /// Level 5: Validate table cell structure.
    /// Every hp:tc must have:
    ///   - cellAddr (child element OR tc attributes with colAddr/rowAddr)
    ///   - cellSz (child element with width and height)
    ///   - cellMargin (child element with left, right, top, bottom)
    ///   - subList with ≥1 hp:p
    /// </summary>
    private void ValidateTableStructure(List<ValidationError> errors)
    {
        foreach (var section in _doc.Sections)
        {
            int tblIdx = 0;
            foreach (var tbl in section.Tables)
            {
                tblIdx++;
                int trIdx = 0;
                foreach (var tr in tbl.Elements(HwpxNs.Hp + "tr"))
                {
                    trIdx++;
                    int tcIdx = 0;
                    foreach (var tc in tr.Elements(HwpxNs.Hp + "tc"))
                    {
                        tcIdx++;
                        var basePath = $"/section[{section.Index + 1}]/tbl[{tblIdx}]/tr[{trIdx}]/tc[{tcIdx}]";
                        var part = section.EntryPath;

                        // Check cellAddr (dual-format: child element OR tc attributes — require BOTH colAddr AND rowAddr)
                        var cellAddrChild = tc.Element(HwpxNs.Hp + "cellAddr");
                        var hasAttrAddr = tc.Attribute("colAddr") != null && tc.Attribute("rowAddr") != null;  // require BOTH
                        if (cellAddrChild != null)
                        {
                            // child element form: must have both attributes
                            if (cellAddrChild.Attribute("colAddr") == null || cellAddrChild.Attribute("rowAddr") == null)
                                errors.Add(new ValidationError(
                                    "table_celladdr_incomplete",
                                    "cellAddr element missing 'colAddr' or 'rowAddr' attribute",
                                    basePath, part));
                        }
                        else if (!hasAttrAddr)
                        {
                            errors.Add(new ValidationError(
                                "table_missing_celladdr",
                                "Table cell missing cellAddr (no child element and no colAddr/rowAddr attributes)",
                                basePath,
                                part));
                        }

                        // Check cellSz
                        var cellSz = tc.Element(HwpxNs.Hp + "cellSz");
                        if (cellSz == null)
                        {
                            errors.Add(new ValidationError(
                                "table_missing_cellsz",
                                "Table cell missing cellSz element (width and height)",
                                basePath,
                                part));
                        }
                        else
                        {
                            if (cellSz.Attribute("width") == null)
                                errors.Add(new ValidationError(
                                    "table_cellsz_no_width",
                                    "cellSz element missing 'width' attribute",
                                    basePath,
                                    part));
                            if (cellSz.Attribute("height") == null)
                                errors.Add(new ValidationError(
                                    "table_cellsz_no_height",
                                    "cellSz element missing 'height' attribute",
                                    basePath,
                                    part));
                        }

                        // Check cellMargin
                        var cellMargin = tc.Element(HwpxNs.Hp + "cellMargin");
                        if (cellMargin == null)
                        {
                            errors.Add(new ValidationError(
                                "table_missing_cellmargin",
                                "Table cell missing cellMargin element (left, right, top, bottom)",
                                basePath,
                                part));
                        }
                        else
                        {
                            foreach (var side in new[] { "left", "right", "top", "bottom" })
                            {
                                if (cellMargin.Attribute(side) == null)
                                    errors.Add(new ValidationError(
                                        "table_cellmargin_incomplete",
                                        $"cellMargin element missing '{side}' attribute",
                                        basePath,
                                        part));
                            }
                        }

                        // Check subList with ≥1 <hp:p>
                        var subList = tc.Element(HwpxNs.Hp + "subList");
                        if (subList == null)
                        {
                            errors.Add(new ValidationError(
                                "table_missing_sublist",
                                "Table cell missing subList element (must contain ≥1 paragraph)",
                                basePath,
                                part));
                        }
                        else
                        {
                            var paraCount = subList.Elements(HwpxNs.Hp + "p").Count();
                            if (paraCount == 0)
                            {
                                errors.Add(new ValidationError(
                                    "table_empty_sublist",
                                    "Table cell subList contains no paragraphs (must have ≥1 hp:p)",
                                    basePath,
                                    part));
                            }
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Level 6: Verify that required namespace declarations exist in respective root elements.
    /// - Hs namespace in section roots
    /// - Hp namespace in section roots
    /// - Hh namespace in header.xml root
    /// </summary>
    private void ValidateNamespaceDeclarations(List<ValidationError> errors)
    {
        // Check section roots for Hs and Hp namespaces
        foreach (var section in _doc.Sections)
        {
            var root = section.Root;
            var declaredNamespaces = root.Attributes()
                .Where(a => a.IsNamespaceDeclaration)
                .Select(a => a.Value)
                .ToHashSet();

            // Also check the namespace of the root element itself
            var rootNs = root.Name.Namespace.NamespaceName;

            if (!declaredNamespaces.Contains(HwpxNs.Hs.NamespaceName)
                && rootNs != HwpxNs.Hs.NamespaceName)
            {
                errors.Add(new ValidationError(
                    "ns_missing",
                    $"Section {section.Index + 1} root missing Hs namespace declaration " +
                    $"({HwpxNs.Hs.NamespaceName})",
                    $"/section[{section.Index + 1}]",
                    section.EntryPath));
            }

            // Check for Hp namespace (used by child elements)
            var hasHpInTree = root.Descendants()
                .Any(e => e.Name.Namespace == HwpxNs.Hp);
            if (hasHpInTree
                && !declaredNamespaces.Contains(HwpxNs.Hp.NamespaceName)
                && rootNs != HwpxNs.Hp.NamespaceName)
            {
                errors.Add(new ValidationError(
                    "ns_missing",
                    $"Section {section.Index + 1} root missing Hp namespace declaration " +
                    $"({HwpxNs.Hp.NamespaceName}) — elements use this namespace",
                    $"/section[{section.Index + 1}]",
                    section.EntryPath));
            }
        }

        // Check header.xml for Hh namespace
        if (_doc.Header?.Root != null)
        {
            var headerRoot = _doc.Header.Root;
            var declaredNamespaces = headerRoot.Attributes()
                .Where(a => a.IsNamespaceDeclaration)
                .Select(a => a.Value)
                .ToHashSet();

            var rootNs = headerRoot.Name.Namespace.NamespaceName;

            if (!declaredNamespaces.Contains(HwpxNs.Hh.NamespaceName)
                && rootNs != HwpxNs.Hh.NamespaceName)
            {
                var hasHhInTree = headerRoot.Descendants()
                    .Any(e => e.Name.Namespace == HwpxNs.Hh);
                if (hasHhInTree)
                {
                    errors.Add(new ValidationError(
                        "ns_missing",
                        $"header.xml root missing Hh namespace declaration " +
                        $"({HwpxNs.Hh.NamespaceName}) — elements use this namespace",
                        "/header",
                        _doc.HeaderEntryPath ?? "Contents/header.xml"));
                }
            }
        }
    }

    // Plan 94: BinData integrity (merged from ViewAsIssues Level 7)
    private void ValidateBinDataIntegrity(List<ValidationError> errors)
    {
        var referenced = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sec in _doc.Sections)
            foreach (var el in sec.Root.Descendants())
            {
                var binRef = el.Attribute("binaryItemIDRef")?.Value;
                if (binRef != null) referenced.Add(binRef);
            }

        foreach (var manifestRef in GetManifestBinDataIds())
            referenced.Add(manifestRef);

        var actual = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in _doc.Archive.Entries)
            if (entry.FullName.Contains("BinData/", StringComparison.OrdinalIgnoreCase)
                && !entry.FullName.EndsWith("/", StringComparison.Ordinal))
                actual.Add(System.IO.Path.GetFileNameWithoutExtension(entry.FullName));

        foreach (var missing in referenced.Except(actual))
            errors.Add(new ValidationError("bindata_missing",
                $"Referenced binary '{missing}' not found in archive",
                "/BinData", null));

        foreach (var orphan in actual.Except(referenced))
            errors.Add(new ValidationError("bindata_orphan",
                $"Orphan binary '{orphan}' not referenced by any element",
                "/BinData", null));
    }

    private IEnumerable<string> GetManifestBinDataIds()
    {
        if (_doc.ManifestDoc?.Root == null) yield break;

        foreach (var item in _doc.ManifestDoc.Descendants())
        {
            var href = item.Attribute("href")?.Value;
            if (href == null || !href.StartsWith("BinData/", StringComparison.OrdinalIgnoreCase))
                continue;

            var id = item.Attribute("id")?.Value;
            if (!string.IsNullOrWhiteSpace(id))
            {
                yield return id;
                continue;
            }

            var fileId = System.IO.Path.GetFileNameWithoutExtension(href);
            if (!string.IsNullOrWhiteSpace(fileId))
                yield return fileId;
        }
    }

    // Plan 94: Field pair consistency (merged from ViewAsIssues Level 8)
    private void ValidateFieldPairs(List<ValidationError> errors)
    {
        foreach (var sec in _doc.Sections)
        {
            var begins = sec.Root.Descendants(HwpxNs.Hp + "fieldBegin").Count();
            var ends = sec.Root.Descendants(HwpxNs.Hp + "fieldEnd").Count();
            if (begins != ends)
                errors.Add(new ValidationError("field_pair_mismatch",
                    $"Section {sec.Index + 1}: {begins} fieldBegin vs {ends} fieldEnd",
                    $"/section[{sec.Index + 1}]", null));
        }
    }

    // Plan 94: Section count consistency (merged from ViewAsIssues Level 9)
    private void ValidateSectionCount(List<ValidationError> errors)
    {
        if (_doc.ManifestDoc == null) return;
        var manifestSectionCount = _doc.ManifestDoc.Descendants()
            .Count(e => e.Attribute("media-type")?.Value?.Contains("section") ?? false);
        if (manifestSectionCount > 0 && manifestSectionCount != _doc.Sections.Count)
        {
            errors.Add(new ValidationError("package_section_mismatch",
                $"Manifest declares {manifestSectionCount} sections but {_doc.Sections.Count} loaded",
                "/Contents/content.hpf", null));
        }
    }
}
