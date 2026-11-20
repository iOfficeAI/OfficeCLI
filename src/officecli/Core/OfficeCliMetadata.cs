// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using AP = DocumentFormat.OpenXml.ExtendedProperties;
using CP = DocumentFormat.OpenXml.CustomProperties;
using VT = DocumentFormat.OpenXml.VariantTypes;

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
    // require lossy mangling of semver, and no major Office UI surfaces it.
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
    /// instead of the standard <c>/docProps/core.xml</c> Office writes.
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
    /// Whether this invocation may write the identification / audit metadata
    /// at all. Off via <c>officecli config stamp false</c>, or for a single
    /// invocation via <c>OFFICECLI_NO_STAMP=1</c>; the environment variable
    /// wins, so a caller can opt out without mutating the user's config file.
    /// Default is on — a missing config file or an unreadable one never
    /// changes what earlier versions wrote.
    ///
    /// Deliberately re-read per call instead of cached in a static: a resident
    /// server outlives the config it loaded, and <c>officecli config stamp
    /// false</c> in another terminal has to apply to the resident's next save,
    /// not to its next restart. The read is one small file, the same cost
    /// <see cref="CliLogger.Enabled"/> pays per log line.
    /// </summary>
    internal static bool StampingEnabled
    {
        get
        {
            // IsTruthySafe and not `== "1"`: an unrecognized value must leave
            // stamping on rather than throw from a save path, and the accepted
            // spellings (1/true/yes/on) are a superset of the sibling flags'.
            if (ParseHelpers.IsTruthySafe(Environment.GetEnvironmentVariable("OFFICECLI_NO_STAMP")))
                return false;
            try { return UpdateChecker.LoadConfig().Stamp ?? true; }
            catch { return true; }
        }
    }

    /// <summary>
    /// Stamp a freshly-created document as authored by OfficeCLI. Writes
    /// <c>docProps/core.xml</c> (Creator, Created, LastModifiedBy, Modified),
    /// <c>docProps/app.xml</c> (Application = "OfficeCLI/&lt;version&gt;", no AppVersion),
    /// and <c>docProps/custom.xml</c> (OfficeCLI.Version, OfficeCLI.LastModified).
    ///
    /// Only invoked from <see cref="BlankDocCreator"/> on initial creation.
    /// For edits to existing documents, use <see cref="StampOnSave"/> which
    /// only updates the audit trail in <c>custom.xml</c> and leaves
    /// <c>app.xml</c>'s &lt;Application&gt; untouched.
    ///
    /// A no-op when <see cref="StampingEnabled"/> is off: nothing is written,
    /// and no core/app/custom part, relationship or content-type override is
    /// created either.
    /// </summary>
    public static void StampOnCreate(OpenXmlPackage doc)
    {
        if (!StampingEnabled) return;

        var nowUtc = DateTime.UtcNow;
        WriteCoreProperties(doc, nowUtc);

        var part = ExtendedPropertiesHandler.GetOrCreateExtendedPart(doc);
        if (part != null)
        {
            part.Properties ??= new AP.Properties();
            (part.Properties.Application ??= new AP.Application()).Text = AppName;
            part.Properties.Save();
        }

        WriteCustomProperties(doc, nowUtc);
    }

    /// <summary>
    /// Stamp an audit trail on every save of an existing document. Writes
    /// only <c>docProps/custom.xml</c> with OfficeCLI.Version and
    /// OfficeCLI.LastModified. Does NOT touch &lt;Application&gt; in app.xml
    /// (preserving the original authoring tool's identity) nor core.xml
    /// (avoiding LastModifiedBy clobbering of the real human author).
    ///
    /// A no-op when <see cref="StampingEnabled"/> is off. Properties an
    /// earlier run already wrote are left in place — the switch stops new
    /// writes, it does not strip existing metadata.
    /// </summary>
    public static void StampOnSave(OpenXmlPackage doc)
    {
        if (!StampingEnabled) return;
        WriteCustomProperties(doc, DateTime.UtcNow);
    }

    private const string CustomPropsFmtId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    private static CustomFilePropertiesPart? GetOrCreateCustomPart(OpenXmlPackage doc) => doc switch
    {
        WordprocessingDocument w => w.CustomFilePropertiesPart ?? w.AddCustomFilePropertiesPart(),
        SpreadsheetDocument s => s.CustomFilePropertiesPart ?? s.AddCustomFilePropertiesPart(),
        PresentationDocument p => p.CustomFilePropertiesPart ?? p.AddCustomFilePropertiesPart(),
        _ => null
    };

    /// <summary>
    /// Write/update OfficeCLI custom properties in <c>docProps/custom.xml</c>.
    /// Properties used: <c>OfficeCLI.Version</c>, <c>OfficeCLI.LastModified</c>.
    /// Existing properties (from any author) are preserved verbatim; only
    /// the two OfficeCLI-owned keys are upserted, and <c>pid</c> values are
    /// renumbered into the contiguous 2..N range the OOXML schema requires.
    /// </summary>
    private static void WriteCustomProperties(OpenXmlPackage doc, DateTime nowUtc)
    {
        var part = GetOrCreateCustomPart(doc);
        if (part == null) return;

        part.Properties ??= new CP.Properties();
        var props = part.Properties;

        void Upsert(string name, string value)
        {
            CP.CustomDocumentProperty? existing = null;
            foreach (var el in props.Elements<CP.CustomDocumentProperty>())
            {
                if (string.Equals(el.Name?.Value, name, StringComparison.Ordinal))
                {
                    existing = el;
                    break;
                }
            }
            if (existing != null)
            {
                existing.RemoveAllChildren();
                existing.AppendChild(new VT.VTLPWSTR(value));
                return;
            }
            var added = new CP.CustomDocumentProperty
            {
                FormatId = CustomPropsFmtId,
                Name = name,
            };
            added.AppendChild(new VT.VTLPWSTR(value));
            props.AppendChild(added);
        }

        Upsert("OfficeCLI.Version", ResolveVersion());
        Upsert("OfficeCLI.LastModified", nowUtc.ToString("yyyy-MM-ddTHH:mm:ssZ"));

        // OOXML requires pid to be a contiguous sequence starting at 2.
        // Renumber every CustomDocumentProperty in document order so the
        // schema stays valid regardless of prior authors' numbering.
        int pid = 2;
        foreach (var el in props.Elements<CP.CustomDocumentProperty>())
        {
            el.PropertyId = pid++;
        }

        props.Save();
    }
}
