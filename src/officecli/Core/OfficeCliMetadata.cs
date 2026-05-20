// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using AP = DocumentFormat.OpenXml.ExtendedProperties;

namespace OfficeCli.Core;

/// <summary>
/// Stamps OOXML packages with OfficeCLI identification (app.xml + core.xml).
/// </summary>
internal static class OfficeCliMetadata
{
    public const string ProductName = "OfficeCLI";

    // Application string follows the convention "<Product>/<Version>"
    // so the version is visible everywhere Application is surfaced (Windows
    // Word's Advanced Properties → Statistics, audit tools, file inspectors).
    // We deliberately omit ap:AppVersion: its OOXML "X.YYYY" format would
    // require lossy mangling of semver, no major Office UI surfaces it, and
    // POI also skips it.
    private static readonly string _appName = $"{ProductName}/{ResolveVersion()}";

    /// <summary>String written to <c>ap:Application</c>, e.g. "OfficeCLI/1.0.58".</summary>
    public static string AppName => _appName;

    /// <summary>Bare product name, written to <c>dc:creator</c> and <c>cp:lastModifiedBy</c>.</summary>
    public static string CreatorName => ProductName;

    private static string ResolveVersion()
    {
        var asm = typeof(OfficeCliMetadata).Assembly;
        var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                   ?? asm.GetName().Version?.ToString()
                   ?? "0.0.0";
        var plus = info.IndexOf('+');
        return plus > 0 ? info[..plus] : info;
    }

    private const string CpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private const string DcNs = "http://purl.org/dc/elements/1.1/";
    private const string DctermsNs = "http://purl.org/dc/terms/";
    private const string XsiNs = "http://www.w3.org/2001/XMLSchema-instance";

    private static CoreFilePropertiesPart? GetOrCreateCorePart(OpenXmlPackage doc) => doc switch
    {
        WordprocessingDocument w => w.CoreFilePropertiesPart ?? w.AddCoreFilePropertiesPart(),
        SpreadsheetDocument s => s.CoreFilePropertiesPart ?? s.AddCoreFilePropertiesPart(),
        PresentationDocument p => p.CoreFilePropertiesPart ?? p.AddCoreFilePropertiesPart(),
        _ => null
    };

    /// <summary>
    /// Marshal core properties directly to the CoreFilePropertiesPart stream.
    /// We bypass <see cref="OpenXmlPackage.PackageProperties"/> on purpose: that
    /// path delegates to <c>System.IO.Packaging.Package.PackageProperties</c>,
    /// which on .NET stores props in a non-canonical
    /// <c>/package/services/metadata/core-properties/&lt;guid&gt;.psmdcp</c> blob
    /// instead of the standard <c>/docProps/core.xml</c> Office and POI write.
    ///
    /// Read-modify-write semantics: every existing element (with its
    /// attributes) is preserved verbatim — including non-standard fields
    /// other producers (Pages / Keynote / WPS) occasionally add — and only the
    /// four OfficeCLI-relevant fields are upserted.
    /// </summary>
    private static void WriteCoreProperties(OpenXmlPackage doc, DateTime nowUtc)
    {
        var part = GetOrCreateCorePart(doc);
        if (part == null) return;

        XElement root;
        try
        {
            using var rs = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
            if (rs.Length > 0)
            {
                var loaded = XDocument.Load(rs).Root;
                root = loaded ?? new XElement(XName.Get("coreProperties", CpNs));
            }
            else
            {
                root = new XElement(XName.Get("coreProperties", CpNs));
            }
        }
        catch
        {
            root = new XElement(XName.Get("coreProperties", CpNs));
        }

        void Upsert(string ns, string local, string value, bool withW3CDTF)
        {
            var name = XName.Get(local, ns);
            var el = root.Element(name);
            if (el == null)
            {
                el = new XElement(name, value);
                if (withW3CDTF)
                    el.SetAttributeValue(XName.Get("type", XsiNs), "dcterms:W3CDTF");
                root.Add(el);
            }
            else
            {
                el.Value = value;
                if (withW3CDTF && el.Attribute(XName.Get("type", XsiNs)) == null)
                    el.SetAttributeValue(XName.Get("type", XsiNs), "dcterms:W3CDTF");
            }
        }

        var iso = nowUtc.ToString("yyyy-MM-ddTHH:mm:ssZ");
        Upsert(DcNs, "creator", CreatorName, withW3CDTF: false);
        Upsert(DctermsNs, "created", iso, withW3CDTF: true);
        Upsert(CpNs, "lastModifiedBy", CreatorName, withW3CDTF: false);
        Upsert(DctermsNs, "modified", iso, withW3CDTF: true);

        // Ensure idiomatic prefixes on the root for the standard four
        // namespaces (Office writes these as cp/dc/dcterms/xsi). XDocument
        // emits each child's namespace as default if no prefix is bound, so
        // pin the prefixes explicitly.
        SetXmlnsIfMissing(root, "cp", CpNs);
        SetXmlnsIfMissing(root, "dc", DcNs);
        SetXmlnsIfMissing(root, "dcterms", DctermsNs);
        SetXmlnsIfMissing(root, "xsi", XsiNs);

        using var ws = part.GetStream(FileMode.Create, FileAccess.Write);
        var settings = new XmlWriterSettings
        {
            Encoding = new System.Text.UTF8Encoding(false),
            OmitXmlDeclaration = false,
        };
        using var xw = XmlWriter.Create(ws, settings);
        xw.WriteStartDocument(true);
        root.WriteTo(xw);
    }

    private static void SetXmlnsIfMissing(XElement el, string prefix, string ns)
    {
        var attrName = XNamespace.Xmlns + prefix;
        if (el.Attribute(attrName) == null)
            el.SetAttributeValue(attrName, ns);
    }

    /// <summary>
    /// Stamp a freshly-created document as authored by OfficeCLI. Writes
    /// <c>docProps/core.xml</c> (Creator, Created, LastModifiedBy, Modified) and
    /// <c>docProps/app.xml</c> (Application = "OfficeCLI/&lt;version&gt;", no AppVersion).
    ///
    /// Only invoked from <see cref="BlankDocCreator"/> on initial creation —
    /// existing documents are left untouched on edit, both to avoid clobbering
    /// foreign tooling's metadata and because read-modify-write of arbitrary
    /// core.xml has unbounded edge cases.
    /// </summary>
    public static void StampOnCreate(OpenXmlPackage doc)
    {
        var nowUtc = DateTime.UtcNow;
        WriteCoreProperties(doc, nowUtc);

        var part = ExtendedPropertiesHandler.GetOrCreateExtendedPart(doc);
        if (part != null)
        {
            part.Properties ??= new AP.Properties();
            (part.Properties.Application ??= new AP.Application()).Text = AppName;
            part.Properties.Save();
        }
    }
}
