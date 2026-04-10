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
    /// </summary>
    internal static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        var parts = path.Split('/', StringSplitOptions.RemoveEmptyEntries);

        foreach (var part in parts)
        {
            var match = Regex.Match(part, @"^(\w+)(?:\[(\d+)\])?$");
            if (!match.Success)
                throw new ArgumentException($"Invalid path segment: '{part}'");

            var name = match.Groups[1].Value;
            int? index = match.Groups[2].Success ? int.Parse(match.Groups[2].Value) : null;
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
        var idx = (segment.Index ?? 1) - 1; // convert 1-indexed to 0-indexed

        XName elementName = name switch
        {
            "p" => HwpxNs.Hp + "p",
            "tbl" => HwpxNs.Hp + "tbl",
            "tr" => HwpxNs.Hp + "tr",
            "tc" => HwpxNs.Hp + "tc",
            "run" => HwpxNs.Hp + "run",
            "img" => HwpxNs.Hp + "img",
            "drawing" => HwpxNs.Hp + "drawing",
            _ => throw new ArgumentException($"Unknown element type: '{name}'")
        };

        var children = parent.Elements(elementName).ToList();
        if (idx < 0 || idx >= children.Count)
            throw new ArgumentException(
                $"{name}[{segment.Index ?? 1}] not found (parent has {children.Count} {name} elements)");

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
        var results = new List<XElement>();
        var trimmed = selector.Trim();

        // Simple element selectors
        if (trimmed == "p")
        {
            foreach (var sec in _doc.Sections)
                results.AddRange(sec.Paragraphs);
            return results;
        }

        if (trimmed == "tbl")
        {
            foreach (var sec in _doc.Sections)
                results.AddRange(sec.Tables);
            return results;
        }

        // Pseudo-selector: p:empty
        if (trimmed == "p:empty")
        {
            foreach (var sec in _doc.Sections)
            {
                foreach (var p in sec.Paragraphs)
                {
                    var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                    if (string.IsNullOrWhiteSpace(text))
                        results.Add(p);
                }
            }
            return results;
        }

        // Pseudo-selector: p:contains(text)
        var containsMatch = Regex.Match(trimmed, @"^p:contains\((.+)\)$");
        if (containsMatch.Success)
        {
            var searchText = containsMatch.Groups[1].Value.Trim('"', '\'');
            foreach (var sec in _doc.Sections)
            {
                foreach (var p in sec.Paragraphs)
                {
                    var text = HwpxKorean.Normalize(ExtractParagraphText(p));
                    if (text.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                        results.Add(p);
                }
            }
            return results;
        }

        // Attribute selector: p[attr=value]
        var attrMatch = Regex.Match(trimmed, @"^(\w+)\[(\w+)=(\w+)\]$");
        if (attrMatch.Success)
        {
            var elemType = attrMatch.Groups[1].Value;
            var attrName = attrMatch.Groups[2].Value;
            var attrValue = attrMatch.Groups[3].Value;

            XName xname = elemType switch
            {
                "p" => HwpxNs.Hp + "p",
                "tbl" => HwpxNs.Hp + "tbl",
                "tr" => HwpxNs.Hp + "tr",
                "tc" => HwpxNs.Hp + "tc",
                _ => throw new ArgumentException($"Unsupported element in selector: '{elemType}'")
            };

            foreach (var sec in _doc.Sections)
            {
                results.AddRange(
                    sec.Root.Descendants(xname)
                        .Where(e => e.Attribute(attrName)?.Value == attrValue));
            }
            return results;
        }

        throw new ArgumentException(
            $"Unsupported selector: '{selector}'. " +
            "Supported: p, tbl, p:empty, p:contains(text), element[attr=value]");
    }
}
