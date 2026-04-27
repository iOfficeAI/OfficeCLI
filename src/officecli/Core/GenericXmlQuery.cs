// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace OfficeCli.Core;

/// <summary>
/// Generic XML-based query engine (Scheme B).
/// Traverses the OpenXML element tree matching by XML local name and attributes.
/// Used as a fallback when the element type is not recognized by handler-specific (Scheme A) logic.
/// </summary>
internal static class GenericXmlQuery
{
    /// <summary>
    /// Query an OpenXML element tree by XML local name and attribute filters.
    /// </summary>
    /// <param name="root">Root element to search within</param>
    /// <param name="elementName">XML local name to match (with optional namespace prefix like "a:ln")</param>
    /// <param name="attributes">Attribute filters: key=value pairs. Value prefixed with "!" means not-equal.</param>
    /// <param name="containsText">If set, only match elements whose InnerText contains this string</param>
    /// <returns>List of matching DocumentNode results</returns>
    public static List<DocumentNode> Query(OpenXmlElement root, string elementName,
        Dictionary<string, string> attributes, string? containsText)
    {
        var results = new List<DocumentNode>();

        // Parse namespace prefix if present (e.g., "a:ln" -> prefix="a", localName="ln")
        string? nsPrefix = null;
        string localName = elementName;
        var colonIdx = elementName.IndexOf(':');
        if (colonIdx > 0)
        {
            nsPrefix = elementName[..colonIdx];
            localName = elementName[(colonIdx + 1)..];
        }

        string? nsUri = null;
        if (nsPrefix != null)
        {
            CommonNamespaces.TryGetValue(nsPrefix, out nsUri);
        }

        Traverse(root, localName, nsUri, attributes, containsText, "", results,
            new Dictionary<string, int>());

        return results;
    }

    private static void Traverse(OpenXmlElement element, string targetLocalName,
        string? targetNsUri, Dictionary<string, string> attributes, string? containsText,
        string parentPath, List<DocumentNode> results,
        Dictionary<string, int> parentCounters)
    {
        var elLocalName = element.LocalName;

        // Build counter key (namespace-qualified to avoid collisions)
        var counterKey = elLocalName;
        if (!parentCounters.ContainsKey(counterKey))
            parentCounters[counterKey] = 0;
        var idx = parentCounters[counterKey];
        parentCounters[counterKey] = idx + 1;

        var currentPath = $"{parentPath}/{elLocalName}[{idx + 1}]";

        // Check if this element matches
        if (MatchesElement(element, targetLocalName, targetNsUri, attributes, containsText))
        {
            results.Add(ElementToNode(element, currentPath));
        }

        // Recurse into children
        var childCounters = new Dictionary<string, int>();
        foreach (var child in element.ChildElements)
        {
            Traverse(child, targetLocalName, targetNsUri, attributes, containsText,
                currentPath, results, childCounters);
        }
    }

    private static bool MatchesElement(OpenXmlElement element, string targetLocalName,
        string? targetNsUri, Dictionary<string, string> attributes, string? containsText)
    {
        // Match local name
        if (!element.LocalName.Equals(targetLocalName, StringComparison.OrdinalIgnoreCase))
            return false;

        // Match namespace if specified
        if (targetNsUri != null && element.NamespaceUri != targetNsUri)
            return false;

        // Match attributes
        foreach (var (key, rawVal) in attributes)
        {
            if (key.StartsWith("__")) continue; // Skip internal pseudo-selectors

            bool negate = rawVal.StartsWith("!");
            var val = negate ? rawVal[1..] : rawVal;

            var actual = GetAttributeValue(element, key);

            bool matches = string.Equals(actual, val, StringComparison.OrdinalIgnoreCase);
            if (negate ? matches : !matches) return false;
        }

        // Match :contains
        if (containsText != null)
        {
            if (!element.InnerText.Contains(containsText, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    /// <summary>
    /// Get attribute value from an element.
    /// First checks direct XML attributes, then checks child elements with a "val" attribute
    /// (common OpenXML pattern: e.g., w:jc w:val="center").
    /// </summary>
    public static string? GetAttributeValue(OpenXmlElement element, string attrName)
    {
        // 1. Check direct XML attributes (by local name)
        foreach (var attr in element.GetAttributes())
        {
            if (attr.LocalName.Equals(attrName, StringComparison.OrdinalIgnoreCase))
                return attr.Value;
        }

        // 2. Check child element val pattern: <child val="..."/>
        foreach (var child in element.ChildElements)
        {
            if (child.LocalName.Equals(attrName, StringComparison.OrdinalIgnoreCase))
            {
                // Look for "val" attribute on this child
                foreach (var attr in child.GetAttributes())
                {
                    if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        return attr.Value;
                }
                // If child exists but has no val, return its InnerText
                if (!string.IsNullOrEmpty(child.InnerText))
                    return child.InnerText;
                return ""; // Child exists but empty
            }
        }

        return null;
    }

    /// <summary>
    /// Convert any OpenXmlElement to a DocumentNode with attributes, text, and optional child recursion.
    /// </summary>
    public static DocumentNode ElementToNode(OpenXmlElement element, string path, int depth = 0)
    {
        var node = new DocumentNode
        {
            Path = path,
            Type = element.LocalName,
            ChildCount = element.ChildElements.Count
        };

        // Set text
        var innerText = element.InnerText;
        if (!string.IsNullOrEmpty(innerText))
        {
            node.Text = innerText.Length > 200 ? innerText[..200] + "..." : innerText;
        }

        // Preview: show XML snippet if no meaningful text
        if (string.IsNullOrEmpty(innerText))
        {
            var outerXml = element.OuterXml;
            node.Preview = outerXml.Length > 200 ? outerXml[..200] + "..." : outerXml;
        }

        // Populate Format with all direct XML attributes
        foreach (var attr in element.GetAttributes())
        {
            node.Format[attr.LocalName] = attr.Value;
        }

        // Also include child element val attributes (common OpenXML pattern)
        foreach (var child in element.ChildElements)
        {
            if (child.ChildElements.Count == 0)
            {
                foreach (var attr in child.GetAttributes())
                {
                    if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                    {
                        node.Format[$"{child.LocalName}"] = attr.Value;
                        break;
                    }
                }
            }
        }

        // Recurse children if depth > 0
        if (depth > 0)
        {
            var typeCounters = new Dictionary<string, int>();
            foreach (var child in element.ChildElements)
            {
                var name = child.LocalName;
                typeCounters.TryGetValue(name, out int idx);
                node.Children.Add(ElementToNode(child, $"{path}/{name}[{idx + 1}]", depth - 1));
                typeCounters[name] = idx + 1;
            }
        }

        return node;
    }

    /// <summary>
    /// Parse a path string like "a/b[1]/c[2]" into segments of (Name, Index).
    /// Index is 1-based. If no index specified, Index is null.
    /// </summary>
    public static List<(string Name, int? Index)> ParsePathSegments(string path)
    {
        var segments = new List<(string Name, int? Index)>();
        foreach (var part in path.Trim('/').Split('/'))
        {
            if (string.IsNullOrEmpty(part)) continue;
            var bracketIdx = part.IndexOf('[');
            if (bracketIdx >= 0)
            {
                var name = PathAliases.Resolve(part[..bracketIdx]);
                var indexStr = part[(bracketIdx + 1)..^1];
                if (!int.TryParse(indexStr, out var idx))
                    throw new ArgumentException($"Invalid path index '{indexStr}' in segment '{part}'. Expected a numeric index.");
                if (idx < 1)
                    throw new ArgumentException($"Invalid path index '{idx}' in segment '{part}'. Index must be >= 1.");
                segments.Add((name, idx));
            }
            else
            {
                segments.Add((PathAliases.Resolve(part), null));
            }
        }
        return segments;
    }

    /// <summary>
    /// Navigate an OpenXML element tree by path segments (localName + optional 1-based index).
    /// Returns null if any segment cannot be resolved.
    /// </summary>
    public static OpenXmlElement? NavigateByPath(OpenXmlElement root, IReadOnlyList<(string Name, int? Index)> segments)
    {
        OpenXmlElement? current = root;
        foreach (var seg in segments)
        {
            if (current == null) return null;
            var children = current.ChildElements
                .Where(e => e.LocalName.Equals(seg.Name, StringComparison.OrdinalIgnoreCase));
            current = seg.Index.HasValue
                ? children.ElementAtOrDefault(seg.Index.Value - 1)
                : children.FirstOrDefault();
        }
        return current;
    }

    /// <summary>
    /// Generic attribute/property setting on an OpenXML element.
    /// Tries: 1) direct XML attribute, 2) existing child element with val attribute,
    /// 3) create new typed child element via SDK (validates against OpenXML schema).
    /// Returns true if the property was set, false if unsupported.
    /// </summary>
    public static bool SetGenericAttribute(OpenXmlElement element, string key, string value)
    {
        // 1. Check direct XML attributes
        foreach (var attr in element.GetAttributes())
        {
            if (attr.LocalName.Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                element.SetAttribute(new OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, value));
                return true;
            }
        }

        // 2. Check existing child element with val pattern
        foreach (var child in element.ChildElements)
        {
            if (child.LocalName.Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                foreach (var attr in child.GetAttributes())
                {
                    if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                    {
                        child.SetAttribute(new OpenXmlAttribute(attr.Prefix, "val", attr.NamespaceUri, value));
                        return true;
                    }
                }
                if (child.ChildElements.Count == 0)
                {
                    child.InnerXml = System.Security.SecurityElement.Escape(value);
                    return true;
                }
                return false;
            }
        }

        // 3. Try creating a new typed child via SDK's type system.
        //    Clone parent (empty), set InnerXml with the new child — SDK will parse it
        //    as a typed element if valid, or OpenXmlUnknownElement if not.
        return TryCreateTypedChild(element, key, value);
    }

    /// <summary>
    /// Try to create and append a typed child element to a parent element.
    /// Uses the SDK's XML parsing to validate: clones the parent (empty), injects
    /// a child XML fragment, checks if the SDK recognizes it as a typed element with Val property.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.UnconditionalSuppressMessage("Trimming", "IL2075",
        Justification = "Probing for a SDK-generated 'Val' property by name. OpenXml typed-element classes are referenced elsewhere and preserved by the trimmer; if trimmed away, the probe simply returns null and the caller falls through.")]
    public static bool TryCreateTypedChild(OpenXmlElement parent, string key, string value)
    {
        var nsUri = parent.NamespaceUri;
        var prefix = parent.Prefix;
        if (string.IsNullOrEmpty(nsUri) || string.IsNullOrEmpty(prefix))
            return false;

        try
        {
            var escapedVal = System.Security.SecurityElement.Escape(value);
            var tempElement = parent.CloneNode(false);
            tempElement.InnerXml = $"<{prefix}:{key} xmlns:{prefix}=\"{nsUri}\" {prefix}:val=\"{escapedVal}\"/>";

            var newChild = tempElement.FirstChild?.CloneNode(true);
            if (newChild == null || newChild is OpenXmlUnknownElement)
                return false;

            // Schema check: only accept "scalar val" typed elements — those that
            // expose a typed Val property on their generated SDK class. Composite
            // types (w:tabs, w:rFonts, w:ind, w:spacing, w:numPr, ...) have no
            // Val property; they'd otherwise accept the fabricated val= as an
            // unknown extension attribute and silently produce invalid XML
            // (e.g. <w:tabs w:val="9360:right"/> instead of nested <w:tab>).
            if (newChild.GetType().GetProperty("Val") == null)
                return false;

            // Schema-aware AddChild rejects elements that don't belong in this
            // parent (e.g. w:snapToGrid in rPr — it's pPr-only). On rejection,
            // return false so the caller can try a different container; do NOT
            // fall back to AppendChild, which bypasses schema and produces
            // invalid XML in the wrong parent.
            if (parent is OpenXmlCompositeElement composite)
            {
                if (!composite.AddChild(newChild, throwOnError: false))
                    return false;
            }
            else
            {
                parent.AppendChild(newChild);
            }

            // Only after AddChild succeeded: remove any older instance the
            // curated reader didn't notice. Doing this earlier would damage
            // existing data on a probe that ultimately fails.
            var existing = parent.ChildElements.FirstOrDefault(e =>
                e.LocalName == key && !ReferenceEquals(e, newChild));
            existing?.Remove();
            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Try to create a new typed child element under a parent, then set multiple properties on it.
    /// Used as the generic fallback for the "add" command when the element type is not recognized
    /// by handler-specific logic. The element is created via SDK's XML parsing (same technique as
    /// TryCreateTypedChild) but without requiring a "val" attribute — properties are set individually
    /// via SetGenericAttribute after creation.
    /// Returns the created element, or null if the SDK does not recognize the type.
    /// </summary>
    public static OpenXmlElement? TryCreateTypedElement(OpenXmlElement parent, string elementName,
        Dictionary<string, string> properties, int? index = null)
    {
        // Support namespace prefix (e.g., "a:solidFill" → prefix="a", localName="solidFill")
        string prefix;
        string localName;
        string nsUri;
        var colonIdx = elementName.IndexOf(':');
        if (colonIdx > 0)
        {
            var nsPrefix = elementName[..colonIdx];
            localName = elementName[(colonIdx + 1)..];
            if (!CommonNamespaces.TryGetValue(nsPrefix, out var resolvedUri))
                return null;
            prefix = nsPrefix;
            nsUri = resolvedUri;
        }
        else
        {
            // Default: use parent's namespace
            nsUri = parent.NamespaceUri;
            prefix = parent.Prefix;
            localName = elementName;
        }

        if (string.IsNullOrEmpty(nsUri) || string.IsNullOrEmpty(prefix))
            return null;

        try
        {
            // Build XML fragment with properties as attributes, so SDK parses them together
            var attrXml = new System.Text.StringBuilder();
            var declaredPrefixes = new HashSet<string> { prefix }; // element prefix already declared
            foreach (var (key, value) in properties)
            {
                var escapedVal = System.Security.SecurityElement.Escape(value);
                // Support namespace-prefixed attributes (e.g., "r:embed", "w:val")
                if (key.Contains(':'))
                {
                    var kColonIdx = key.IndexOf(':');
                    var attrPrefix = key[..kColonIdx];
                    var attrLocal = key[(kColonIdx + 1)..];
                    if (CommonNamespaces.TryGetValue(attrPrefix, out var attrNsUri))
                    {
                        attrXml.Append($" {attrPrefix}:{attrLocal}=\"{escapedVal}\"");
                        if (declaredPrefixes.Add(attrPrefix))
                            attrXml.Append($" xmlns:{attrPrefix}=\"{attrNsUri}\"");
                    }
                }
                else
                {
                    attrXml.Append($" {key}=\"{escapedVal}\"");
                }
            }

            var tempParent = parent.CloneNode(false);
            tempParent.InnerXml = $"<{prefix}:{localName} xmlns:{prefix}=\"{nsUri}\"{attrXml}/>";

            var newChild = tempParent.FirstChild?.CloneNode(true);
            if (newChild == null || newChild is OpenXmlUnknownElement)
                return null;

            // For any properties that weren't set as XML attributes (e.g., child-element val patterns),
            // try SetGenericAttribute as fallback
            foreach (var (key, value) in properties)
            {
                // Skip if already set as XML attribute
                var attrLocal = key.Contains(':') ? key[(key.IndexOf(':') + 1)..] : key;
                if (newChild.GetAttributes().Any(a => a.LocalName.Equals(attrLocal, StringComparison.OrdinalIgnoreCase)))
                    continue;
                SetGenericAttribute(newChild, key, value);
            }

            // Insert: use schema-aware AddChild for correct element ordering,
            // fall back to manual index-based insertion if specified
            if (index.HasValue)
            {
                var children = parent.ChildElements.ToList();
                if (index.Value >= 0 && index.Value < children.Count)
                    children[index.Value].InsertBeforeSelf(newChild);
                else
                    parent.AppendChild(newChild);
            }
            else if (parent is OpenXmlCompositeElement composite)
            {
                // AddChild uses Metadata.Particle.Set() to find correct schema position
                if (!composite.AddChild(newChild, throwOnError: false))
                    parent.AppendChild(newChild); // fallback if schema doesn't define this child
            }
            else
            {
                parent.AppendChild(newChild);
            }

            return newChild;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Parse a CSS-like selector into element name, attributes, and containsText.
    /// Reusable by all handlers for Scheme B fallback.
    /// </summary>
    public static (string element, Dictionary<string, string> attrs, string? containsText) ParseSelector(string selector)
    {
        var attrs = new Dictionary<string, string>();
        string? containsText = null;

        // Extract element name (before any [ or : modifier)
        // Support namespace prefix with colon (e.g., "a:ln"), so find '[' or ':' that starts a pseudo-selector
        var firstBracket = selector.IndexOf('[');
        var pseudoIdx = selector.IndexOf(":contains(", StringComparison.Ordinal);
        var emptyIdx = selector.IndexOf(":empty", StringComparison.Ordinal);
        var noAltIdx = selector.IndexOf(":no-alt", StringComparison.Ordinal);

        var firstMod = selector.Length;
        if (firstBracket >= 0 && firstBracket < firstMod) firstMod = firstBracket;
        if (pseudoIdx >= 0 && pseudoIdx < firstMod) firstMod = pseudoIdx;
        if (emptyIdx >= 0 && emptyIdx < firstMod) firstMod = emptyIdx;
        if (noAltIdx >= 0 && noAltIdx < firstMod) firstMod = noAltIdx;

        var element = selector[..firstMod].Trim();

        // Parse [attr=value] attributes (\\?! handles zsh escaping \! as !)
        foreach (Match m in Regex.Matches(selector, @"\[([\w:]+)(\\?!?=)([^\]]+)\]"))
        {
            var key = m.Groups[1].Value;
            var op = m.Groups[2].Value.Replace("\\", "");
            var val = m.Groups[3].Value.Trim('\'', '"');
            attrs[key] = (op == "!=" ? "!" : "") + val;
        }

        // Parse :contains("text")
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) containsText = containsMatch.Groups[1].Value;

        return (element, attrs, containsText);
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
}
