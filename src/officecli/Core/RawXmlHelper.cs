// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OfficeCli.Core;

/// <summary>
/// Shared helper for raw XML operations (read/write via XPath).
/// This enables AI to perform any OpenXML operation by manipulating XML directly.
/// </summary>
public static class RawXmlHelper
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
        // Convert OpenXml tree to XDocument for XPath support
        var xDoc = XDocument.Parse(rootElement.OuterXml);
        var nsManager = BuildNamespaceManager(xDoc);

        var nodes = xDoc.XPathSelectElements(xpath, nsManager).ToList();
        if (nodes.Count == 0)
        {
            Console.Error.WriteLine($"raw-set: XPath matched no elements: {xpath}");
            Console.Error.WriteLine("Hint: auto-registered namespace prefixes: " +
                string.Join(", ", CommonNamespaces.Keys.Order()) +
                ". No xmlns declarations needed in --xml fragments.");
            return 0;
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

        return affected;
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
        var validator = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        return validator.Validate(package)
            .Select(e => new ValidationError(
                e.ErrorType.ToString(),
                e.Description,
                e.Path?.XPath,
                e.Part?.Uri.ToString()))
            .ToList();
    }

    private static void TryAddNamespace(XmlNamespaceManager nsManager, string prefix, string uri)
    {
        if (string.IsNullOrEmpty(nsManager.LookupNamespace(prefix)))
        {
            nsManager.AddNamespace(prefix, uri);
        }
    }
}
