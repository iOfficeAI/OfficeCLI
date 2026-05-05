// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    internal record PathSegment(string Name, int? Index);

    /// <summary>
    /// Parse a path string into segments.
    /// "/section[1]/p[3]" → [("section", 1), ("p", 3)]
    /// "/section[1]/p[last]" → [("section", 1), ("p", -1)]
    /// -1 is a sentinel value meaning "last element".
    /// </summary>
    internal static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        var parts = path.Split('/', StringSplitOptions.RemoveEmptyEntries);

        foreach (var part in parts)
        {
            var match = Regex.Match(part, @"^(\w+)(?:\[(\d+|last)\])?$");
            if (!match.Success)
                throw new ArgumentException($"Invalid path segment: '{part}'");

            var name = match.Groups[1].Value;
            int? index;
            if (match.Groups[2].Success)
                index = match.Groups[2].Value == "last" ? -1 : int.Parse(match.Groups[2].Value);
            else
                index = null;
            segments.Add(new PathSegment(name, index));
        }

        return segments;
    }

    /// <summary>
    /// Resolve a path to the target XElement.
    /// Supports cross-section navigation: /section[2]/p[3]
    /// </summary>
    internal XElement ResolvePath(string path)
    {
        if (string.IsNullOrEmpty(path) || path == "/")
            throw new ArgumentException("Cannot resolve root path to element");

        var segments = ParsePath(path);
        return ResolveSegments(segments);
    }

    private XElement ResolveSegments(List<PathSegment> segments)
    {
        if (segments.Count == 0)
            throw new ArgumentException("Empty path");

        var first = segments[0];

        // Determine which section to use
        HwpxSection section;
        int segmentStart;

        if (first.Name.Equals("section", StringComparison.OrdinalIgnoreCase))
        {
            var secIdx = (first.Index ?? 1) - 1;
            if (secIdx < 0 || secIdx >= _doc.Sections.Count)
                throw new ArgumentException(
                    $"Section {secIdx + 1} not found (document has {_doc.Sections.Count} sections)");
            section = _doc.Sections[secIdx];
            segmentStart = 1; // skip section segment
        }
        else if (first.Name.Equals("header", StringComparison.OrdinalIgnoreCase))
        {
            return ResolveHeaderPath(segments);
        }
        else
        {
            // No section prefix → use primary section
            section = _doc.PrimarySection;
            segmentStart = 0;
        }

        if (segmentStart >= segments.Count)
        {
            // Path is just "/section[N]" — return section root
            return section.Root;
        }

        // Resolve within section
        XElement current = section.Root;
        for (int i = segmentStart; i < segments.Count; i++)
        {
            var seg = segments[i];
            current = ResolveChildElement(current, seg);
        }

        return current;
    }

    private XElement ResolveChildElement(XElement parent, PathSegment segment)
    {
        var name = segment.Name.ToLowerInvariant();

        XName elementName = name switch
        {
            "p" => HwpxNs.Hp + "p",
            "tbl" => HwpxNs.Hp + "tbl",
            "tr" => HwpxNs.Hp + "tr",
            "tc" => HwpxNs.Hp + "tc",
            "run" => HwpxNs.Hp + "run",
            "pic" or "picture" or "image" => HwpxNs.Hp + "pic",
            "img" => HwpxNs.Hp + "img",
            "drawing" => HwpxNs.Hp + "drawing",
            "sublist" => HwpxNs.Hp + "subList",
            _ => throw new ArgumentException($"Unknown element type: '{name}'")
        };

        // When parent is <hp:tc>, transparently navigate through <hp:subList>
        // so paths like /tbl[1]/tr[1]/tc[1]/p[1] work without explicit subList
        var searchParent = parent;
        if (parent.Name.LocalName == "tc" && (name == "p" || name == "run"))
        {
            searchParent = parent.Element(HwpxNs.Hp + "subList") ?? parent;
        }

        // When looking for tbl inside a p, search descendants (tbl is inside p > run > tbl)
        var children = (parent.Name.LocalName == "p" && name == "tbl")
            ? parent.Descendants(elementName).ToList()
            : searchParent.Elements(elementName).ToList();

        // Resolve index: -1 = last, null = first (1), positive = 1-based
        int idx;
        if (segment.Index == -1)
            idx = children.Count - 1;
        else
            idx = (segment.Index ?? 1) - 1;

        if (idx < 0 || idx >= children.Count)
        {
            var label = segment.Index == -1 ? "last" : (segment.Index ?? 1).ToString();
            throw new ArgumentException(
                $"{name}[{label}] not found (parent has {children.Count} {name} elements)");
        }

        return children[idx];
    }

    private XElement ResolveHeaderPath(List<PathSegment> segments)
    {
        if (_doc.Header?.Root == null)
            throw new ArgumentException("Document has no header.xml");

        if (segments.Count == 1)
            return _doc.Header.Root;

        var second = segments[1];
        var name = second.Name.ToLowerInvariant();

        // Navigate header.xml structure
        XName elementName = name switch
        {
            "charpr" or "charproperty" => HwpxNs.Hh + "charPr",
            "parapr" or "paraproperty" => HwpxNs.Hh + "paraPr",
            "style" => HwpxNs.Hh + "style",
            "borderfill" => HwpxNs.Hh + "borderFill",
            _ => throw new ArgumentException($"Unknown header element: '{name}'")
        };

        if (second.Index.HasValue)
        {
            // Find by ID attribute (header elements use id= not positional index)
            var element = _doc.Header.Root.Descendants(elementName)
                .FirstOrDefault(e => e.Attribute("id")?.Value == second.Index.Value.ToString());
            if (element == null)
                throw new ArgumentException($"{name} with id={second.Index.Value} not found");
            return element;
        }

        // Return container
        var container = name switch
        {
            "charpr" or "charproperty" => HwpxNs.Hh + "charProperties",
            "parapr" or "paraproperty" => HwpxNs.Hh + "paraProperties",
            "style" => HwpxNs.Hh + "styles",
            "borderfill" => HwpxNs.Hh + "borderFills",
            _ => throw new ArgumentException($"Unknown header container: '{name}'")
        };

        return _doc.Header.Root.Descendants(container).FirstOrDefault()
            ?? throw new ArgumentException($"No {name} container found");
    }

    /// <summary>
    /// Build a path string for a given XElement by walking up the tree.
    /// </summary>
    internal string BuildPath(XElement element)
    {
        var parts = new Stack<string>();
        var current = element;

        while (current != null && current.Parent != null)
        {
            var localName = current.Name.LocalName;
            var ns = current.Name.Namespace;

            // Skip subList in path building — users navigate tc/p directly
            if (localName == "subList")
            {
                current = current.Parent;
                continue;
            }

            if (ns == HwpxNs.Hs && localName == "sec")
            {
                // Find section index
                var secIdx = _doc.Sections.FindIndex(s => s.Root == current);
                if (secIdx >= 0)
                    parts.Push($"section[{secIdx + 1}]");
                break;
            }

            // Count siblings of same type to determine index — always emit [N] for consistent string-equality
            var siblings = current.Parent.Elements(current.Name).ToList();
            var idx = siblings.IndexOf(current) + 1;  // 1-based
            parts.Push($"{MapElementToPathName(localName)}[{idx}]");

            current = current.Parent;
        }

        return "/" + string.Join("/", parts);
    }

    private static string MapElementToPathName(string localName) => localName switch
    {
        "p" => "p",
        "tbl" => "tbl",
        "tr" => "tr",
        "tc" => "tc",
        "run" => "run",
        "t" => "t",
        "img" => "img",
        "drawing" => "drawing",
        "subList" => "subList",
        _ => localName
    };

    /// <summary>
    /// Parse and execute a CSS-like selector against the document.
    /// Supported selectors:
    ///   "p" — all paragraphs
    ///   "tbl" — all tables
    ///   "p:empty" — empty paragraphs
    ///   "p:contains(text)" — paragraphs containing text
    ///   "tbl > tr > tc" — table cells (descendant combinator)
    ///   "p[styleIDRef=2]" — attribute selector
    /// </summary>
    internal List<XElement> ExecuteSelector(string selector)
    {
        var trimmed = selector.Trim();

        // Child combinator: "tbl > tr > tc", "p > run"
        if (trimmed.Contains(" > "))
        {
            var parts = trimmed.Split(" > ", StringSplitOptions.TrimEntries);
            var current = GetAllElements(parts[0]);
            for (int i = 1; i < parts.Length; i++)
                current = current.SelectMany(parent => FilterChildren(parent, parts[i])).ToList();
            return current;
        }

        // element:pseudo or element[attr op value]
        var selectorMatch = Regex.Match(trimmed, @"^(\w+)(?::(\w+)(?:\((.+)\))?)?(?:\[(.+)\])?$");
        if (!selectorMatch.Success)
            throw new ArgumentException($"Unsupported selector: '{selector}'. " +
                "Supported: p, tbl, tr, tc, run, picture, p:empty, element:contains(text), element:has(child), element[attr=value], element[attr!=value], element[attr~=text], parent > child");

        var elemType = selectorMatch.Groups[1].Value;
        var pseudo = selectorMatch.Groups[2].Value;
        var pseudoArg = selectorMatch.Groups[3].Value;
        var attrExpr = selectorMatch.Groups[4].Value;

        // Get base elements
        var results = GetAllElements(elemType);

        // Apply pseudo-selector
        if (!string.IsNullOrEmpty(pseudo))
        {
            results = pseudo switch
            {
                "empty" => results.Where(e => string.IsNullOrWhiteSpace(GetElementText(e))).ToList(),
                "contains" => results.Where(e =>
                    GetElementText(e).Contains(pseudoArg.Trim('"', '\''), StringComparison.OrdinalIgnoreCase)).ToList(),
                "has" => results.Where(e =>
                    e.Descendants(ResolveXName(pseudoArg)).Any()).ToList(),
                "first" => results.Take(1).ToList(),
                "last" => results.TakeLast(1).ToList(),
                _ => throw new ArgumentException($"Unsupported pseudo-selector: ':{pseudo}'")
            };
        }

        // Apply attribute filter
        if (!string.IsNullOrEmpty(attrExpr))
            results = ApplyAttributeFilter(results, attrExpr);

        return results;
    }

    private List<XElement> GetAllElements(string elemType)
    {
        var results = new List<XElement>();
        // Strip pseudo/attr for base element resolution
        var baseType = Regex.Replace(elemType, @"[:[].*$", "");

        foreach (var sec in _doc.Sections)
        {
            var xname = ResolveXName(baseType);
            if (baseType == "p")
                results.AddRange(sec.Paragraphs);
            else if (baseType == "tbl")
                results.AddRange(sec.Tables);
            else
                results.AddRange(sec.Root.Descendants(xname));
        }
        return results;
    }

    private static XName ResolveXName(string elemType) => elemType switch
    {
        "p" => HwpxNs.Hp + "p",
        "tbl" => HwpxNs.Hp + "tbl",
        "tr" => HwpxNs.Hp + "tr",
        "tc" => HwpxNs.Hp + "tc",
        "run" => HwpxNs.Hp + "run",
        "pic" or "picture" or "image" => HwpxNs.Hp + "pic",
        "img" => HwpxNs.Hp + "img",
        "equation" => HwpxNs.Hp + "equation",
        "shape" => HwpxNs.Hp + "shapeObject",
        "field" or "fieldBegin" or "formfield" => HwpxNs.Hp + "fieldBegin",
        "bookmark" => HwpxNs.Hp + "bookmark",
        "ctrl" => HwpxNs.Hp + "ctrl",
        _ => HwpxNs.Hp + elemType
    };

    private string GetElementText(XElement e) => e.Name.LocalName switch
    {
        "p" => HwpxKorean.Normalize(ExtractParagraphText(e)),
        "tc" => ExtractCellText(e),
        "run" => string.Join("", e.Elements(HwpxNs.Hp + "t").Select(t => t.Value)),
        _ => e.Value
    };

    private List<XElement> FilterChildren(XElement parent, string childSelector)
    {
        var childMatch = Regex.Match(childSelector, @"^(\w+)(?:\[(.+)\])?$");
        if (!childMatch.Success) return new();
        var childType = childMatch.Groups[1].Value;
        var xname = ResolveXName(childType);
        var children = parent.Elements(xname).ToList();
        if (!string.IsNullOrEmpty(childMatch.Groups[2].Value))
            children = ApplyAttributeFilter(children, childMatch.Groups[2].Value);
        return children;
    }

    private List<XElement> ApplyAttributeFilter(List<XElement> elements, string attrExpr)
    {
        // Operators: =, !=, ~= (contains), >=, <=
        var opMatch = Regex.Match(attrExpr, @"^(\w+)(~=|!=|>=|<=|=)(.+)$");
        if (!opMatch.Success) return elements;

        var attrName = opMatch.Groups[1].Value;
        var op = opMatch.Groups[2].Value;
        var attrValue = opMatch.Groups[3].Value.Trim('"', '\'');

        return elements.Where(e =>
        {
            var actual = ResolveVirtualAttribute(e, attrName);
            if (actual == null) return op == "!="; // null != anything is true
            return op switch
            {
                "=" => actual.Equals(attrValue, StringComparison.OrdinalIgnoreCase),
                "!=" => !actual.Equals(attrValue, StringComparison.OrdinalIgnoreCase),
                "~=" => actual.Contains(attrValue, StringComparison.OrdinalIgnoreCase),
                ">=" => int.TryParse(actual, out var av) && int.TryParse(attrValue, out var tv) && av >= tv,
                "<=" => int.TryParse(actual, out var av2) && int.TryParse(attrValue, out var tv2) && av2 <= tv2,
                _ => false
            };
        }).ToList();
    }

    /// <summary>Resolve virtual attributes (computed on-the-fly) or real XML attributes.</summary>
    private string? ResolveVirtualAttribute(XElement e, string attrName)
    {
        // Virtual attributes
        switch (attrName)
        {
            case "text":
                return GetElementText(e);
            case "bold":
            {
                var charPrId = e.Attribute("charPrIDRef")?.Value ?? e.Elements(HwpxNs.Hp + "run").FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
                var charPr = charPrId != null ? FindCharPr(charPrId) : null;
                return charPr?.Element(HwpxNs.Hh + "bold") != null ? "true" : "false";
            }
            case "italic":
            {
                var charPrId = e.Attribute("charPrIDRef")?.Value ?? e.Elements(HwpxNs.Hp + "run").FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
                var charPr = charPrId != null ? FindCharPr(charPrId) : null;
                return charPr?.Element(HwpxNs.Hh + "italic") != null ? "true" : "false";
            }
            case "fontsize":
            {
                var charPrId = e.Attribute("charPrIDRef")?.Value ?? e.Elements(HwpxNs.Hp + "run").FirstOrDefault()?.Attribute("charPrIDRef")?.Value;
                var charPr = charPrId != null ? FindCharPr(charPrId) : null;
                var height = (int?)charPr?.Attribute("height");
                return height.HasValue ? (height.Value / 100).ToString() : null;
            }
            case "colSpan":
                return e.Name.LocalName == "tc" ? GetCellAddr(e).ColSpan.ToString() : e.Attribute("colSpan")?.Value;
            case "rowSpan":
                return e.Name.LocalName == "tc" ? GetCellAddr(e).RowSpan.ToString() : e.Attribute("rowSpan")?.Value;
            case "heading":
            {
                if (e.Name.LocalName != "p") return null;
                var info = GetParagraphStyleInfo(e);
                return info.HeadingLevel;
            }
        }

        // Real XML attribute
        return e.Attribute(attrName)?.Value;
    }
}
