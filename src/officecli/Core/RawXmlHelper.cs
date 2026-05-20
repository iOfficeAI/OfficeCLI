// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Compression;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
// OOXML0001: PackageExtensions.GetPackage() is marked [Experimental] but it
// is the only public path to the underlying IPackage. The previous reflection
// workaround tripped trim analysis (IL2026/IL2075/IL2060) and was fragile —
// IPackage/IPackagePart are themselves public interfaces with public methods,
// so once GetPackage() returns we are entirely on the public surface.
#pragma warning disable OOXML0001
using DocumentFormat.OpenXml.Experimental;
#pragma warning restore OOXML0001

namespace OfficeCli.Core;

/// <summary>
/// Shared helper for raw XML operations (read/write via XPath).
/// This enables AI to perform any OpenXML operation by manipulating XML directly.
/// </summary>
internal static class RawXmlHelper
{
    /// <summary>
    /// Perform a raw XML operation on a document part's root element.
    /// </summary>
    /// <param name="rootElement">The OpenXml root element (e.g. Document, Worksheet, Slide)</param>
    /// <param name="xpath">XPath expression to locate target element(s)</param>
    /// <param name="action">Operation: append, prepend, insertbefore, insertafter, replace, remove, setattr</param>
    /// <param name="xml">XML fragment for append/prepend/insert/replace, or attr=value for setattr</param>
    /// <returns>Number of elements affected</returns>
    public static int Execute(OpenXmlPartRootElement rootElement, string xpath, string action, string? xml)
    {
        var (xDoc, affected) = ExecuteOnXmlString(rootElement.OuterXml, xpath, action, xml);
        if (affected == 0) return 0;
        // Write modified XML back to the OpenXml element.
        // Propagate namespace declarations from root to direct child elements,
        // so that each child's ToString() produces self-contained XML with
        // all necessary namespace bindings (otherwise inherited namespaces are lost).
        var rootNsAttrs = xDoc.Root!.Attributes()
            .Where(a => a.IsNamespaceDeclaration).ToList();
        foreach (var child in xDoc.Root.Elements())
        {
            foreach (var nsAttr in rootNsAttrs)
            {
                if (child.Attribute(nsAttr.Name) == null)
                    child.SetAttributeValue(nsAttr.Name, nsAttr.Value);
            }
        }
        rootElement.InnerXml = string.Concat(xDoc.Root.Nodes().Select(n => n.ToString()));
        // The InnerXml setter restores inner content but does NOT touch root
        // attributes — so non-xmlns attrs like `mc:Ignorable` carried by the
        // replacement root would be silently lost on round-trip. Copy them
        // through. Skip xmlns declarations (SDK manages those via its typed
        // prefix table). Need the prefix for namespaced attrs so SDK renders
        // them as `<prefix>:<localname>` rather than re-prefixing under the
        // SDK's auto-generated alias.
        foreach (var attr in xDoc.Root.Attributes())
        {
            if (attr.IsNamespaceDeclaration) continue;
            var ns = attr.Name.NamespaceName;
            // Resolve the prefix as the source XML had it (walk parent's
            // in-scope namespace declarations). XLinq's XAttribute.Name has
            // no Prefix property, only the parent element does.
            var prefix = string.IsNullOrEmpty(ns)
                ? string.Empty
                : (xDoc.Root.GetPrefixOfNamespace(ns) ?? string.Empty);
            var openXmlAttr = new DocumentFormat.OpenXml.OpenXmlAttribute(
                prefix, attr.Name.LocalName, ns, attr.Value);
            rootElement.SetAttribute(openXmlAttr);
        }
        return affected;
    }

    /// <summary>
    /// Apply a raw XML operation directly on a part's stream (no SDK typed
    /// root needed). Used for arbitrary XML parts addressed by zip URI —
    /// sheet1.xml, footnotes.xml, customXml/item1.xml, etc.
    /// </summary>
    public static int Execute(OpenXmlPart part, string xpath, string action, string? xml)
    {
        var sourceXml = ReadPartXml(part);
        var (xDoc, affected) = ExecuteOnXmlString(sourceXml, xpath, action, xml);
        if (affected == 0) return 0;
        WritePartXml(part, xDoc.ToString(SaveOptions.DisableFormatting));
        return affected;
    }

    private static (XDocument xDoc, int affected) ExecuteOnXmlString(
        string sourceXml, string xpath, string action, string? xml)
    {
        var xDoc = XDocument.Parse(sourceXml);
        var nsManager = BuildNamespaceManager(xDoc);

        var nodes = xDoc.XPathSelectElements(xpath, nsManager).ToList();
        if (nodes.Count == 0)
        {
            Console.Error.WriteLine($"raw-set: XPath matched no elements: {xpath}");
            Console.Error.WriteLine("Hint: auto-registered namespace prefixes: " +
                string.Join(", ", CommonNamespaces.Keys.Order()) +
                ". No xmlns declarations needed in --xml fragments.");
            return (xDoc, 0);
        }

        int affected = 0;

        foreach (var node in nodes)
        {
            switch (action.ToLowerInvariant())
            {
                case "append":
                    if (xml == null) throw new ArgumentException("--xml is required for append");
                    var appendFragment = ParseFragment(xml, xDoc);
                    foreach (var el in appendFragment)
                        node.Add(el);
                    affected++;
                    break;

                case "prepend":
                    if (xml == null) throw new ArgumentException("--xml is required for prepend");
                    var prependFragment = ParseFragment(xml, xDoc);
                    foreach (var el in prependFragment.AsEnumerable().Reverse())
                        node.AddFirst(el);
                    affected++;
                    break;

                case "insertbefore" or "before":
                    if (xml == null) throw new ArgumentException("--xml is required for insertbefore");
                    var beforeFragment = ParseFragment(xml, xDoc);
                    foreach (var el in beforeFragment.AsEnumerable().Reverse())
                        node.AddBeforeSelf(el);
                    affected++;
                    break;

                case "insertafter" or "after":
                    if (xml == null) throw new ArgumentException("--xml is required for insertafter");
                    var afterFragment = ParseFragment(xml, xDoc);
                    foreach (var el in afterFragment)
                        node.AddAfterSelf(el);
                    affected++;
                    break;

                case "replace":
                    if (xml == null) throw new ArgumentException("--xml is required for replace");
                    var replaceFragment = ParseFragment(xml, xDoc);
                    node.ReplaceWith(replaceFragment.ToArray());
                    affected++;
                    break;

                case "remove" or "delete":
                    node.Remove();
                    affected++;
                    break;

                case "setattr":
                    if (xml == null) throw new ArgumentException("--xml is required for setattr (format: name=value)");
                    var eqIdx = xml.IndexOf('=');
                    if (eqIdx <= 0) throw new ArgumentException("setattr format: name=value");
                    var attrName = xml[..eqIdx];
                    var attrValue = xml[(eqIdx + 1)..];

                    // Handle namespaced attributes (e.g. w:val)
                    var colonIdx = attrName.IndexOf(':');
                    if (colonIdx > 0)
                    {
                        var prefix = attrName[..colonIdx];
                        var localName = attrName[(colonIdx + 1)..];
                        var ns = nsManager.LookupNamespace(prefix);
                        if (ns != null)
                            node.SetAttributeValue(XName.Get(localName, ns), attrValue);
                        else
                            node.SetAttributeValue(attrName, attrValue);
                    }
                    else
                    {
                        node.SetAttributeValue(attrName, attrValue);
                    }
                    affected++;
                    break;

                default:
                    throw new ArgumentException($"Unknown action: {action}. Supported: append, prepend, insertbefore, insertafter, replace, remove, setattr");
            }
        }

        return (xDoc, affected);
    }

    private static readonly Dictionary<string, string> CommonNamespaces = new()
    {
        ["w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ["r"] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ["a"] = "http://schemas.openxmlformats.org/drawingml/2006/main",
        ["p"] = "http://schemas.openxmlformats.org/presentationml/2006/main",
        ["x"] = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ["wp"] = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ["c"] = "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ["xdr"] = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        ["wps"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ["wp14"] = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ["v"] = "urn:schemas-microsoft-com:vml",
    };

    private static List<XElement> ParseFragment(string xml, XDocument contextDoc)
    {
        // Collect namespace declarations from the context document
        var nsDict = new Dictionary<string, string>(CommonNamespaces);
        string? defaultNs = null;

        if (contextDoc.Root != null)
        {
            // Inherit the default namespace from the document root so that
            // unprefixed elements (e.g. <mergeCells>) are parsed into the
            // correct namespace (e.g. spreadsheetml) instead of empty namespace.
            var rootNsName = contextDoc.Root.Name.NamespaceName;
            if (!string.IsNullOrEmpty(rootNsName))
                defaultNs = rootNsName;

            foreach (var attr in contextDoc.Root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                var prefix = attr.Name.Namespace == XNamespace.Xmlns ? attr.Name.LocalName : "";
                if (!string.IsNullOrEmpty(prefix))
                    nsDict[prefix] = attr.Value;
            }
        }

        var prefixedNs = string.Join(" ", nsDict.Select(kv => $"xmlns:{kv.Key}=\"{kv.Value}\""));
        var defaultNsDecl = !string.IsNullOrEmpty(defaultNs) ? $"xmlns=\"{defaultNs}\"" : "";
        var wrappedXml = $"<_root {defaultNsDecl} {prefixedNs}>{xml}</_root>";
        var parsed = XDocument.Parse(wrappedXml);
        return parsed.Root!.Elements().ToList();
    }

    private static XmlNamespaceManager BuildNamespaceManager(XDocument xDoc)
    {
        var nsManager = new XmlNamespaceManager(new NameTable());

        if (xDoc.Root != null)
        {
            foreach (var attr in xDoc.Root.Attributes().Where(a => a.IsNamespaceDeclaration))
            {
                var prefix = attr.Name.LocalName;
                if (attr.Name.Namespace == XNamespace.Xmlns)
                {
                    nsManager.AddNamespace(prefix, attr.Value);
                }
                else if (attr.Name == "xmlns")
                {
                    // Default namespace — assign a usable prefix
                    nsManager.AddNamespace("default", attr.Value);
                }
            }
        }

        // Ensure common OpenXML namespaces are available
        TryAddNamespace(nsManager, "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        TryAddNamespace(nsManager, "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        TryAddNamespace(nsManager, "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        TryAddNamespace(nsManager, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
        TryAddNamespace(nsManager, "x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        TryAddNamespace(nsManager, "wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
        TryAddNamespace(nsManager, "mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        TryAddNamespace(nsManager, "c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        TryAddNamespace(nsManager, "xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");

        return nsManager;
    }

    /// <summary>
    /// Validate an OpenXmlPackage and return structured errors.
    /// </summary>
    public static List<ValidationError> ValidateDocument(OpenXmlPackage package)
    {
        var errors = PreflightXmlParts(package);
        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        // BUG-R6-08: documents containing w:numPicBullet can trip an NRE
        // inside SDK validation when one of its child accessors hits a
        // null. Materialise per-error with try/catch so a single problem
        // entry doesn't bring the whole `validate` command down. Surface
        // the exception as a synthetic ValidationError instead of
        // bubbling out as a process-level crash.
        IEnumerable<DocumentFormat.OpenXml.Validation.ValidationErrorInfo> raw;
        try
        {
            raw = validator.Validate(package);
        }
        catch (Exception ex)
        {
            errors.Add(new ValidationError(
                "ValidatorException",
                $"Validator threw before producing results: {ex.GetType().Name}: {ex.Message}",
                null, null));
            return errors;
        }
        // The IEnumerable is lazy — iterate with try/catch so one bad
        // error entry does not abort the rest.
        using var enumerator = raw.GetEnumerator();
        while (true)
        {
            DocumentFormat.OpenXml.Validation.ValidationErrorInfo? e = null;
            try
            {
                if (!enumerator.MoveNext()) break;
                e = enumerator.Current;
            }
            catch (NullReferenceException nre)
            {
                errors.Add(new ValidationError(
                    "ValidatorNullReference",
                    $"SDK validator hit a null while inspecting next error: {nre.Message}",
                    null, null));
                continue;
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"SDK validator threw while inspecting next error: {ex.GetType().Name}: {ex.Message}",
                    null, null));
                continue;
            }
            if (e == null) continue;
            try
            {
                errors.Add(new ValidationError(
                    e.ErrorType.ToString(),
                    e.Description,
                    e.Path?.XPath,
                    e.Part?.Uri.ToString()));
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "ValidatorException",
                    $"Failed to materialise validation error: {ex.GetType().Name}: {ex.Message}",
                    null, null));
            }
        }
        return errors;
    }

    /// <summary>
    /// Walk every XML part in the package and try to parse it. OpenXmlValidator
    /// silently skips parts that can't be read, so a deck whose presentation.xml
    /// or any slide part is malformed XML (well-formed zip, broken XML inside)
    /// otherwise reports "validate passed: no errors found" — a dangerous
    /// false negative that hides obvious corruption. Surface a synthetic
    /// MalformedXml validation error for each part that fails to parse so the
    /// caller sees real failure instead of clean success.
    /// </summary>
    private static List<ValidationError> PreflightXmlParts(OpenXmlPackage package)
    {
        var errors = new List<ValidationError>();
        foreach (var part in package.GetAllParts())
        {
            // Only inspect XML parts; binary streams (images, fonts, embedded
            // OLE) don't participate in schema validation. The naive
            // ContentType.Contains("xml") screen matches every OOXML content
            // type via the "openxmlformats" substring (e.g.
            // "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            // for an OLE-embedded .docx), so a perfectly valid binary
            // OOXML embedding was mis-classified as malformed XML on every
            // validate run. RFC 7303 / OOXML restrict actual XML content types
            // to the "+xml" structured-suffix or the plain text/xml family;
            // gate the preflight on those instead.
            var ct = part.ContentType ?? "";
            if (!IsXmlContentType(ct)) continue;
            try
            {
                using var s = part.GetStream(FileMode.Open, FileAccess.Read);
                using var r = XmlReader.Create(s, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreWhitespace = false,
                });
                while (r.Read()) { /* drain */ }
            }
            catch (Exception ex)
            {
                errors.Add(new ValidationError(
                    "MalformedXml",
                    $"Part contains malformed XML and was skipped by schema validation: {ex.GetType().Name}: {ex.Message}",
                    null,
                    part.Uri.ToString()));
            }
        }
        return errors;
    }

    private static bool IsXmlContentType(string ct)
    {
        if (string.IsNullOrEmpty(ct)) return false;
        // Drop parameters (e.g. "; charset=utf-8") and lowercase the type.
        var semi = ct.IndexOf(';');
        var bare = (semi >= 0 ? ct.Substring(0, semi) : ct).Trim().ToLowerInvariant();
        // Structured-suffix form: anything/anything+xml — covers every OOXML
        // schema part (...wordprocessingml.document.main+xml,
        // ...presentationml.slide+xml, ...drawingml.theme+xml, etc.).
        if (bare.EndsWith("+xml")) return true;
        // Generic XML content types.
        return bare == "application/xml" || bare == "text/xml";
    }

    private static void TryAddNamespace(XmlNamespaceManager nsManager, string prefix, string uri)
    {
        if (string.IsNullOrEmpty(nsManager.LookupNamespace(prefix)))
        {
            nsManager.AddNamespace(prefix, uri);
        }
    }

    // ==================== Zip-URI part lookup ====================
    //
    // Rule: any partPath ending in `.xml` is treated as a literal zip-internal
    // URI (e.g. `/xl/worksheets/sheet1.xml`, `/word/footnotes.xml`,
    // `/ppt/slides/slide1.xml`). We walk the entire part tree of the package
    // and match against `OpenXmlPart.Uri.OriginalString`.
    //
    // This supersedes the per-handler hand-curated alias tables, which could
    // never be complete (only covered global parts like /xl/workbook.xml).
    // Semantic paths (`/Sheet1`, `/workbook`, `/document`, `/header[1]`) still
    // route through the handler's own switch — only `.xml`-suffixed inputs
    // hit this lookup.

    /// <summary>
    /// Returns true if `partPath` should be resolved as a literal zip-internal
    /// URI rather than a semantic short name. Trims surrounding whitespace
    /// and discards any URI fragment (`#...`) or query (`?...`) suffix so
    /// `/xl/workbook.xml#frag` and `/xl/workbook.xml?x=1` both classify as
    /// zip-URI inputs (rather than silently falling through to the
    /// semantic-path "Available: ..." dispatcher).
    ///
    /// Accepts `.xml`, `.rels` (relationship parts), and the literal
    /// `[Content_Types].xml` package manifest. The first two are normal OPC
    /// parts; `[Content_Types].xml` is package metadata reachable only
    /// through a separate code path.
    /// </summary>
    public static bool IsZipUriPath(string partPath)
    {
        if (partPath == null) return false;
        var s = StripUriSuffixes(partPath.AsSpan().Trim());
        return s.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            || s.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
    }

    private static ReadOnlySpan<char> StripUriSuffixes(ReadOnlySpan<char> s)
    {
        var cut = s.IndexOfAny('#', '?');
        return cut >= 0 ? s[..cut] : s;
    }

    /// <summary>
    /// Walk the entire part tree of a package and return the part whose
    /// `Uri.OriginalString` matches `partPath` (with leading-slash
    /// normalization). Returns null if no part matches.
    /// </summary>
    /// <summary>
    /// Resolve a zip-URI path through the SDK's own underlying IPackage.
    /// Used as the primary fallback after the typed-OpenXmlPart graph,
    /// because it shares the SDK's file handle and so works correctly when
    /// the file is held open editable (resident mode) where a fresh BCL
    /// Package.Open would fail with a FileShare conflict. Returns null if
    /// the part does not exist.
    /// </summary>
    private static string? TryReadViaSdkPackage(OpenXmlPackage package, string partPath)
    {
        try
        {
            var clean = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
            var target = clean.StartsWith('/') ? clean : "/" + clean;
            Uri uri;
            try { uri = new Uri(target, UriKind.Relative); } catch { return null; }

#pragma warning disable OOXML0001
            var pkg = package.GetPackage();
#pragma warning restore OOXML0001
            if (!pkg.PartExists(uri)) return null;
            var part = pkg.GetPart(uri);

            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();
            var stripped = StripXmlProlog(content);
            if (stripped.Length == 0)
                throw new InvalidDataException(
                    $"Part '{target}' contains no root element (only an XML " +
                    $"declaration, whitespace, or BOM).");
            return stripped;
        }
        catch (InvalidDataException) { throw; }
        catch
        {
            return null;
        }
    }

    public static OpenXmlPart? FindPartByZipUri(OpenXmlPackage package, string partPath)
    {
        // Trim surrounding whitespace, discard fragment/query, and normalize
        // leading slash. Fragments and query strings are not part of OPC
        // URIs; users may inadvertently type them and we should resolve
        // against the bare part path.
        partPath = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
        var target = partPath.StartsWith('/') ? partPath : "/" + partPath;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        return Walk(package);

        OpenXmlPart? Walk(OpenXmlPartContainer container)
        {
            foreach (var rel in container.Parts)
            {
                var p = rel.OpenXmlPart;
                var uri = p.Uri.OriginalString;
                if (!seen.Add(uri)) continue;
                if (string.Equals(uri, target, StringComparison.OrdinalIgnoreCase))
                    return p;
                var nested = Walk(p);
                if (nested != null) return nested;
            }
            return null;
        }
    }

    /// <summary>
    /// Read a part's XML content as a string. Prefer the typed
    /// <see cref="OpenXmlPartRootElement"/> when available (preserves
    /// canonical SDK serialization); fall back to the underlying stream for
    /// untyped XML parts (e.g. CustomXml).
    ///
    /// Output omits the &lt;?xml ?&gt; prolog uniformly so that:
    ///   raw /workbook              (semantic path, typed OuterXml)
    ///   raw /xl/workbook.xml       (zip URI, typed OuterXml)
    ///   raw /customXml/item1.xml   (zip URI, untyped stream)
    /// all produce element-only output. The semantic short-name path has
    /// always done this; this method extends the convention to untyped
    /// parts so zip-URI calls don't randomly include the prolog depending
    /// on whether the SDK strongly-typed the target part.
    /// </summary>
    /// <summary>
    /// Resolve a zip-URI path to its content. Tries the OpenXmlPart graph
    /// first (typed parts — preserves SDK-canonical serialization for parts
    /// that have a strongly-typed root); falls back to the underlying OPC
    /// package (covers relationship parts `.rels` and any XML part the
    /// SDK doesn't surface as a typed OpenXmlPart).
    ///
    /// Returns null if no part matches; throws InvalidDataException if the
    /// part exists but contains no root element.
    /// </summary>
    public static string? TryReadByZipUri(OpenXmlPackage package, string? filePath, string partPath)
    {
        // Typed-part path first (preserves SDK-canonical serialization for
        // strongly-typed parts).
        var typed = FindPartByZipUri(package, partPath);
        if (typed != null) return ReadPartXml(typed);

        // Then: SDK's own underlying IPackage via reflection. This sees
        // every .rels part the SDK is managing AND coexists with the SDK
        // file handle (no second-handle FileShare conflict — important for
        // resident mode where the file is open editable and a fresh
        // BCL Package.Open would fail).
        var sdkResult = TryReadViaSdkPackage(package, partPath);
        if (sdkResult != null) return sdkResult;

        if (filePath == null) return null;

        // Special case: `[Content_Types].xml` is the OPC package manifest,
        // not a part. System.IO.Packaging.Package does not expose it; read
        // it as a literal zip entry.
        var trimmed = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
        if (trimmed.TrimStart('/').Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                using var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
                var entry = zip.Entries.FirstOrDefault(e =>
                    e.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));
                if (entry == null) return null;
                using var es = entry.Open();
                using var er = new StreamReader(es);
                var ec = er.ReadToEnd();
                var es2 = StripXmlProlog(ec);
                return es2.Length == 0
                    ? throw new InvalidDataException("[Content_Types].xml is empty.")
                    : es2;
            }
            catch (Exception ex) when (ex is not InvalidDataException)
            {
                return null;
            }
        }

        var clean = StripUriSuffixes(partPath.AsSpan().Trim()).ToString();
        var target = clean.StartsWith('/') ? clean : "/" + clean;
        Uri uri;
        try { uri = new Uri(target, UriKind.Relative); } catch { return null; }

        Package bclPkg;
        try
        {
            // FileShare.ReadWrite so we coexist with the SDK's existing handle.
            // Package internally opens the underlying stream with
            // FileAccess.Read here.
            bclPkg = Package.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }
        catch
        {
            // SDK opened with FileShare.None (e.g. editable mode on Windows),
            // or the file is gone — give up cleanly.
            return null;
        }

        try
        {
            if (!bclPkg.PartExists(uri)) return null;
            var part = bclPkg.GetPart(uri);
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream);
            var content = reader.ReadToEnd();
            var stripped = StripXmlProlog(content);
            if (stripped.Length == 0)
                throw new InvalidDataException(
                    $"Part '{target}' contains no root element (only an XML " +
                    $"declaration, whitespace, or BOM).");
            return stripped;
        }
        finally
        {
            bclPkg.Close();
        }
    }

    public static string ReadPartXml(OpenXmlPart part)
    {
        if (part.RootElement is OpenXmlPartRootElement root && root != null)
            return root.OuterXml;
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var reader = new StreamReader(stream);
        var content = reader.ReadToEnd();
        var stripped = StripXmlProlog(content);
        if (stripped.Length == 0)
            throw new InvalidDataException(
                $"Part '{part.Uri.OriginalString}' contains no root element " +
                $"(only an XML declaration, whitespace, or BOM). The package " +
                $"may be corrupt; investigate before treating output as data.");
        return stripped;
    }

    private static string StripXmlProlog(string xml)
    {
        var s = xml.AsSpan().TrimStart();
        // Loop: handle multiple stacked prologs / BOMs (defensive — input may
        // be byte-concatenated from upstream tools or a corrupted package).
        while (s.Length > 0)
        {
            // BOM (U+FEFF). StreamReader normally consumes it but we may be
            // reading a re-encoded inner segment.
            if (s[0] == '﻿') { s = s[1..].TrimStart(); continue; }

            // XML declaration: per spec must be `<?xml` followed by whitespace
            // or `?>`. Crucially must NOT match other PIs whose target starts
            // with `xml` (e.g. `<?xml-stylesheet ...?>`), which is a legal
            // processing instruction we must preserve.
            if (s.Length >= 6
                && s[0] == '<' && s[1] == '?'
                && s[2] == 'x' && s[3] == 'm' && s[4] == 'l'
                && (s[5] == ' ' || s[5] == '\t' || s[5] == '\n' || s[5] == '\r'
                    || (s[5] == '?' && s.Length >= 7 && s[6] == '>')))
            {
                var end = s.IndexOf("?>", StringComparison.Ordinal);
                if (end < 0) break;
                s = s[(end + 2)..].TrimStart();
                continue;
            }
            break;
        }
        return s.ToString();
    }

    /// <summary>
    /// Write XML content into a part's stream, replacing prior contents.
    /// </summary>
    public static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }
}
