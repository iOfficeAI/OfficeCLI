// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // BUG-TESTER fuzz-2: bound regex match time on user-supplied find patterns to
    // prevent catastrophic-backtracking DoS (e.g. "(a+)+b" against long inputs).
    private static readonly TimeSpan FindRegexMatchTimeout = TimeSpan.FromSeconds(5);

    private static bool IsTruthy(string? value) =>
        ParseHelpers.IsTruthy(value);

    private static bool IsValidBooleanString(string? value) =>
        ParseHelpers.IsValidBooleanString(value);

    /// <summary>
    /// Normalize cell[R,C] shorthand to tr[R]/tc[C] in paths.
    /// E.g. /slide[1]/table[1]/cell[2,3] → /slide[1]/table[1]/tr[2]/tc[3]
    /// Also handles trailing segments: /slide[1]/table[1]/cell[2,3]/txBody → /slide[1]/table[1]/tr[2]/tc[3]/txBody
    /// </summary>
    /// <summary>
    /// CONSISTENCY(path-stability): the per-handler path-pattern regexes are mostly
    /// case-sensitive. DOCX folds case via ToLowerInvariant on every segment name
    /// (Navigation.cs); we mirror that here by lowercasing the alphabetic LocalName
    /// portion of every `<name>[index]` segment so `/SLIDE[1]/SHAPE[2]` is treated
    /// identically to `/slide[1]/shape[2]` and routes through the structured matchers
    /// instead of falling through to the raw-XML default.
    /// </summary>
    private static string NormalizePptxPathSegmentCasing(string path)
    {
        if (string.IsNullOrEmpty(path) || path == "/") return path;
        // Lowercase only the LocalName before '[' or '/' or end-of-segment. Preserve
        // bracketed identifiers (placeholder[Title 1]), attribute selectors (@role=ROLE),
        // and named arguments verbatim — only the leading element-name token is folded.
        return Regex.Replace(path, @"(?<=^|/)([A-Za-z][A-Za-z0-9]*)",
            m => m.Value.ToLowerInvariant());
    }

    private static string NormalizeCellPath(string path)
    {
        // Reject malformed segment separators that previously slipped past
        // the regex matchers and ended up exposing raw OOXML local names
        // (e.g. `Get("/slide[1]/")` returned type=sld, `Get("//slide[1]")`
        // returned sld). DOCX already rejects these forms; bring PPTX/XLSX
        // up to parity with an explicit error rather than silent leakage.
        if (path.Length > 1 && path != "/" && path.EndsWith("/"))
            throw new ArgumentException($"Invalid path '{path}': trailing '/' is not allowed.");
        if (path.StartsWith("//"))
            throw new ArgumentException($"Invalid path '{path}': leading '//' is not allowed.");
        if (path.Contains("//"))
            throw new ArgumentException($"Invalid path '{path}': empty path segment ('//') is not allowed.");
        return Regex.Replace(path, @"cell\[(\d+),\s*(\d+)\]", m => $"tr[{m.Groups[1].Value}]/tc[{m.Groups[2].Value}]");
    }

    /// <summary>
    /// Resolve InsertPosition (After/Before anchor path) to a 0-based int? index for PPT.
    /// Anchor path can be full (/slide[1]/shape[@id=X]) or short (shape[@id=X]).
    /// </summary>
    /// <summary>Sentinel value for find: anchor resolution.</summary>
    private const int FindAnchorIndex = -99999;

    private int? ResolveAnchorPosition(string parentPath, InsertPosition? position)
    {
        if (position == null) return null;
        if (position.Index.HasValue) return position.Index;

        var anchorPath = position.After ?? position.Before!;

        // Catch bare attribute selector without element wrapper, e.g. @id=XXX instead of shape[@id=XXX]
        if (Regex.IsMatch(anchorPath, @"^@(\w+)=(.+)$"))
            throw new ArgumentException($"Invalid anchor path \"{anchorPath}\". Did you mean: shape[{anchorPath}]?");

        // Handle find: prefix — text-based anchoring
        if (anchorPath.StartsWith("find:", StringComparison.OrdinalIgnoreCase))
            return FindAnchorIndex;

        // Normalize: if short form, prepend parentPath
        if (!anchorPath.StartsWith("/"))
            anchorPath = parentPath.TrimEnd('/') + "/" + anchorPath;

        // Resolve @id=/@name= in the anchor path
        anchorPath = ResolveIdPath(anchorPath);

        // For slide-level anchors (/slide[N])
        var slideMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]$");
        if (slideMatch.Success)
        {
            var slideIdx = int.Parse(slideMatch.Groups[1].Value) - 1; // 0-based
            var slideCount = GetSlideParts().Count();
            if (slideIdx < 0 || slideIdx >= slideCount)
                throw new ArgumentException($"Anchor slide not found: {anchorPath} (total slides: {slideCount})");
            if (position.After != null)
                return slideIdx + 1 >= slideCount ? null : slideIdx + 1;
            else
                return slideIdx;
        }

        // For element-level anchors (/slide[N]/shape[M], /slide[N]/table[M], etc.)
        var elemMatch = Regex.Match(anchorPath, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\]$");
        if (elemMatch.Success)
        {
            var slideIdx = int.Parse(elemMatch.Groups[1].Value);
            var elemIdx = int.Parse(elemMatch.Groups[3].Value) - 1; // 0-based
            // Validate that the anchor element exists
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Anchor slide not found: {anchorPath} (total slides: {slideParts.Count})");
            var anchorShapeTree = GetSlide(slideParts[slideIdx - 1]).CommonSlideData?.ShapeTree;
            if (anchorShapeTree != null)
            {
                var contentChildren = anchorShapeTree.ChildElements
                    .Where(e => e is not NonVisualGroupShapeProperties && e is not GroupShapeProperties)
                    .ToList();
                if (elemIdx < 0 || elemIdx >= contentChildren.Count)
                    throw new ArgumentException($"Anchor element not found: {anchorPath} (total elements on slide: {contentChildren.Count})");
            }
            if (position.After != null)
                return elemIdx + 1; // InsertAtPosition handles bounds
            else
                return elemIdx;
        }

        throw new ArgumentException($"Cannot resolve anchor path: {anchorPath}");
    }

    /// <summary>
    /// Resolve @id= and @name= attribute selectors in a PPT path to positional indices.
    /// E.g. /slide[1]/shape[@id=5] → /slide[1]/shape[N] where N is the positional index of shape with cNvPr.Id=5.
    /// </summary>
    private string ResolveIdPath(string path)
    {
        // Null/empty paths are a valid "duplicate in place" / "no target"
        // signal from CopyFrom and friends; pass them through untouched so
        // downstream dispatch can interpret the null itself.
        if (path == null) return path!;
        // Quick check: if no [@, nothing to resolve
        if (!path.Contains("[@"))
            return path;

        // Iterate matches left-to-right so we can rewrite the prefix as we go;
        // each successive @id=/@name= resolves relative to whatever group context
        // the earlier (already-rewritten) prefix established.
        var sb = new System.Text.StringBuilder();
        var cursor = 0;
        var rewritten = path;
        var matches = Regex.Matches(path, @"(\w+)\[@(id|name)=([^\]]+)\]");
        foreach (Match m in matches)
        {
            sb.Append(path, cursor, m.Index - cursor);
            var prefix = sb.ToString();

            var elementType = m.Groups[1].Value.ToLowerInvariant();
            var attrName = m.Groups[2].Value.ToLowerInvariant();
            var attrValue = m.Groups[3].Value.Trim('"', '\'', ' ');

            var slideMatch = Regex.Match(prefix, @"/slide\[(\d+)\]");
            if (!slideMatch.Success)
                throw new ArgumentException($"Cannot resolve @{attrName}= outside of a slide context: {path}");
            var slideIdx = int.Parse(slideMatch.Groups[1].Value);

            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide {slideIdx} not found (total: {slideParts.Count})");
            var slidePart = slideParts[slideIdx - 1];
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null)
                throw new ArgumentException($"Slide {slideIdx} has no shape tree");

            // CONSISTENCY(group-id-scope): if the prefix has /group[N] segments
            // after /slide[N], scope the @id=/@name= search inside that nested
            // group's shape tree, not the slide-level shape tree.
            OpenXmlElement scope = shapeTree;
            var groupMatches = Regex.Matches(prefix, @"/group\[(\d+)\]");
            foreach (Match gm in groupMatches)
            {
                var gIdx = int.Parse(gm.Groups[1].Value);
                var groups = scope.Elements<GroupShape>().ToList();
                if (gIdx < 1 || gIdx > groups.Count)
                    throw new ArgumentException($"Group {gIdx} not found in scope (total: {groups.Count})");
                scope = groups[gIdx - 1];
            }

            var positionalIdx = FindElementByAttrInScope(scope, elementType, attrName, attrValue);
            var replacement = $"{m.Groups[1].Value}[{positionalIdx}]";
            sb.Append(replacement);
            cursor = m.Index + m.Length;
        }
        sb.Append(path, cursor, path.Length - cursor);
        return sb.ToString();
    }

    /// <summary>
    /// Find the 1-based positional index of an element within its type group by @id= or @name=.
    /// </summary>
    private static int FindElementByAttr(ShapeTree shapeTree, string elementType, string attrName, string attrValue)
        => FindElementByAttrInScope(shapeTree, elementType, attrName, attrValue);

    /// <summary>
    /// Like <see cref="FindElementByAttr"/> but searches direct children of any
    /// container element (ShapeTree or GroupShape). Used to scope @id=/@name=
    /// lookups inside nested groups.
    /// </summary>
    private static int FindElementByAttrInScope(OpenXmlElement scope, string elementType, string attrName, string attrValue)
    {
        var elements = elementType switch
        {
            "shape" or "textbox" or "title" or "equation" => scope.Elements<Shape>()
                .Select(s => (element: (OpenXmlElement)s, nvPr: s.NonVisualShapeProperties?.NonVisualDrawingProperties)).ToList(),
            "picture" or "pic" or "image" => scope.Elements<Picture>()
                .Select(p => (element: (OpenXmlElement)p, nvPr: p.NonVisualPictureProperties?.NonVisualDrawingProperties)).ToList(),
            "table" => scope.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<Drawing.Table>().Any())
                .Select(gf => (element: (OpenXmlElement)gf, nvPr: gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties)).ToList(),
            "chart" => scope.Elements<GraphicFrame>()
                .Where(gf => gf.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any() || IsExtendedChartFrame(gf))
                .Select(gf => (element: (OpenXmlElement)gf, nvPr: gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties)).ToList(),
            "connector" or "connection" => scope.Elements<ConnectionShape>()
                .Select(c => (element: (OpenXmlElement)c, nvPr: c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties)).ToList(),
            "group" => scope.Elements<GroupShape>()
                .Select(g => (element: (OpenXmlElement)g, nvPr: g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties)).ToList(),
            "video" or "audio" => scope.Elements<Picture>()
                .Select(p => (element: (OpenXmlElement)p, nvPr: p.NonVisualPictureProperties?.NonVisualDrawingProperties)).ToList(),
            _ => throw new ArgumentException($"Unknown element type '{elementType}' for @{attrName}= addressing")
        };

        for (int i = 0; i < elements.Count; i++)
        {
            var nvPr = elements[i].nvPr;
            if (nvPr == null) continue;

            if (attrName == "id" && nvPr.Id?.Value.ToString() == attrValue)
                return i + 1;
            if (attrName == "name" && MatchesShapeName(nvPr.Name?.Value, attrValue))
                return i + 1;
        }

        throw new ArgumentException($"No {elementType} found with @{attrName}={attrValue}");
    }

    /// <summary>
    /// Scan all slides to initialize the global shape ID counter.
    /// Called once on document open (editable mode).
    /// </summary>
    private void InitShapeIdCounter()
    {
        const uint minStartId = 10000;
        _usedShapeIds = new HashSet<uint>();
        uint maxId = minStartId - 1;

        foreach (var slidePart in GetSlideParts())
        {
            var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
            if (shapeTree == null) continue;
            foreach (var nvPr in shapeTree.Descendants<NonVisualDrawingProperties>())
            {
                if (nvPr.Id?.HasValue == true)
                {
                    _usedShapeIds.Add(nvPr.Id.Value);
                    if (nvPr.Id.Value > maxId)
                        maxId = nvPr.Id.Value;
                }
            }
        }

        _nextShapeId = maxId + 1;
        if (_nextShapeId < maxId) // uint overflow
            _nextShapeId = minStartId;
    }

    /// <summary>
    /// Generate a unique deterministic cNvPr.Id across all slides.
    /// Uses global instance counter for reproducible, non-repeating IDs.
    /// </summary>
    private uint GenerateUniqueShapeId(ShapeTree shapeTree)
    {
        const uint minStartId = 10000;
        var startId = _nextShapeId;
        while (true)
        {
            var id = _nextShapeId;
            _nextShapeId++;
            if (_nextShapeId < id) // uint overflow
                _nextShapeId = minStartId;
            if (_usedShapeIds.Add(id))
                return id;
            if (_nextShapeId == startId)
                throw new InvalidOperationException("No available shape ID slots");
        }
    }

    /// <summary>
    /// Get the cNvPr.Id for an element, or null if not available.
    /// Works for Shape, Picture, GraphicFrame, ConnectionShape, GroupShape.
    /// </summary>
    internal static uint? GetCNvPrId(OpenXmlElement element)
    {
        return element switch
        {
            Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
            Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value,
            GraphicFrame gf => gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id?.Value,
            ConnectionShape c => c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
            GroupShape g => g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties?.Id?.Value,
            _ => null
        };
    }

    /// <summary>
    /// Build a path segment using @id= if the element has a cNvPr.Id, otherwise use positional index.
    /// E.g. "shape[@id=5]" or "shape[2]".
    /// </summary>
    internal static string BuildElementPathSegment(string elementType, OpenXmlElement element, int positionalIndex)
    {
        var id = GetCNvPrId(element);
        return id.HasValue
            ? $"{elementType}[@id={id.Value}]"
            : $"{elementType}[{positionalIndex}]";
    }

    /// <summary>
    /// Find existing Transition element or create one, avoiding duplicates with unknown-element transitions.
    /// </summary>
    private static Transition FindOrCreateTransition(Slide slide)
    {
        var typed = slide.GetFirstChild<Transition>();
        if (typed != null) return typed;

        // Check for unknown-element transitions (injected as raw XML to survive SDK serialization)
        var unknown = slide.ChildElements.FirstOrDefault(c => c.LocalName == "transition" && c is not Transition);
        if (unknown != null)
        {
            // Replace with a typed Transition so we can set properties
            var trans = new Transition();
            foreach (var attr in unknown.GetAttributes()) trans.SetAttribute(attr);
            trans.InnerXml = unknown.InnerXml;
            unknown.InsertAfterSelf(trans);
            unknown.Remove();
            return trans;
        }

        return slide.AppendChild(new Transition());
    }

    /// <summary>
    /// Set advanceTime on a slide, handling morph AlternateContent correctly.
    /// </summary>
    internal static void SetAdvanceTime(Slide slide, string value)
    {
        var acMorph = slide.ChildElements.FirstOrDefault(c =>
            c.LocalName == "AlternateContent" && c.InnerXml.Contains("morph"));
        if (acMorph != null)
        {
            // Set advTm directly on transitions inside AlternateContent
            foreach (var trans in acMorph.Descendants().Where(d => d.LocalName == "transition"))
                trans.SetAttribute(new OpenXmlAttribute("", "advTm", null!, value));
        }
        else
        {
            FindOrCreateTransition(slide).AdvanceAfterTime = value;
        }
    }

    /// <summary>
    /// Set advanceOnClick on a slide, handling morph AlternateContent correctly.
    /// </summary>
    internal static void SetAdvanceClick(Slide slide, bool value)
    {
        var acMorph = slide.ChildElements.FirstOrDefault(c =>
            c.LocalName == "AlternateContent" && c.InnerXml.Contains("morph"));
        if (acMorph != null)
        {
            foreach (var trans in acMorph.Descendants().Where(d => d.LocalName == "transition"))
                trans.SetAttribute(new OpenXmlAttribute("", "advClick", null!, value ? "1" : "0"));
        }
        else
        {
            FindOrCreateTransition(slide).AdvanceOnClick = value;
        }
    }

    private static double ParseFontSize(string value) =>
        ParseHelpers.ParseFontSize(value);

    /// <summary>
    /// Read table cell border properties following POI's getBorderWidth/getBorderColor pattern.
    /// Maps a:lnL/lnR/lnT/lnB → border.left, border.right, border.top, border.bottom in Format.
    /// </summary>
    private static void ReadTableCellBorders(Drawing.TableCellProperties tcPr, DocumentNode node)
    {
        ReadBorderLine(tcPr.LeftBorderLineProperties, "border.left", node);
        ReadBorderLine(tcPr.RightBorderLineProperties, "border.right", node);
        ReadBorderLine(tcPr.TopBorderLineProperties, "border.top", node);
        ReadBorderLine(tcPr.BottomBorderLineProperties, "border.bottom", node);
        ReadBorderLine(tcPr.TopLeftToBottomRightBorderLineProperties, "border.tl2br", node);
        ReadBorderLine(tcPr.BottomLeftToTopRightBorderLineProperties, "border.tr2bl", node);
    }

    /// <summary>
    /// Read a single border line's properties (color, width, dash) following POI's pattern:
    /// - Returns nothing if line is null, has NoFill, or lacks SolidFill
    /// - Reads width from w attribute, color from SolidFill, dash from PresetDash
    /// </summary>
    private static void ReadBorderLine(OpenXmlCompositeElement? lineProps, string prefix, DocumentNode node)
    {
        if (lineProps == null) return;
        // POI: if NoFill is set, the border is invisible — skip
        if (lineProps.GetFirstChild<Drawing.NoFill>() != null) return;
        var solidFill = lineProps.GetFirstChild<Drawing.SolidFill>();
        if (solidFill == null) return; // POI: !isSetSolidFill → null

        var color = ReadColorFromFill(solidFill);
        if (color != null) node.Format[$"{prefix}.color"] = color;

        // Width from "w" attribute (EMU) — POI: Units.toPoints(ln.getW())
        var wAttr = lineProps.GetAttributes().FirstOrDefault(a => a.LocalName == "w");
        if (!string.IsNullOrEmpty(wAttr.Value) && long.TryParse(wAttr.Value, out var wEmu) && wEmu > 0)
            node.Format[$"{prefix}.width"] = FormatEmu(wEmu);

        // Dash style from PresetDash — POI: ln.getPrstDash().getVal()
        var dash = lineProps.GetFirstChild<Drawing.PresetDash>();
        if (dash?.Val?.HasValue == true)
            node.Format[$"{prefix}.dash"] = dash.Val.InnerText;

        // Summary key: "1pt solid FF0000" format for convenience
        var parts = new List<string>();
        if (!string.IsNullOrEmpty(wAttr.Value) && long.TryParse(wAttr.Value, out var wEmu2) && wEmu2 > 0)
            parts.Add(FormatEmu(wEmu2));
        if (dash?.Val?.HasValue == true) parts.Add(dash.Val.InnerText!);
        else parts.Add("solid");
        if (color is not null) parts.Add(color);
        if (parts.Count > 0) node.Format[prefix] = string.Join(" ", parts);
    }

    private static string GetShapeText(Shape shape)
    {
        var textBody = shape.TextBody;
        if (textBody == null) return "";

        var sb = new StringBuilder();
        var first = true;
        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            if (!first) sb.Append('\n');
            first = false;
            foreach (var child in para.ChildElements)
            {
                if (child is Drawing.Run run)
                    sb.Append(run.Text?.Text ?? "");
                else if (HasMathContent(child))
                    sb.Append(FormulaParser.ToReadableText(GetMathElement(child)));
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Find all OMML math elements inside a shape's text body.
    /// </summary>
    private static List<OpenXmlElement> FindShapeMathElements(Shape shape)
    {
        var results = new List<OpenXmlElement>();
        var textBody = shape.TextBody;
        if (textBody == null) return results;

        foreach (var para in textBody.Elements<Drawing.Paragraph>())
        {
            foreach (var child in para.ChildElements)
            {
                if (HasMathContent(child))
                    results.Add(GetMathElement(child));
            }
        }
        return results;
    }

    /// <summary>
    /// Check if an element contains math content (a14:m or mc:AlternateContent with math).
    /// </summary>
    private static bool HasMathContent(OpenXmlElement element)
    {
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
            return true;
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            if (element.Descendants().Any(e => e.LocalName == "oMath" || e.LocalName == "oMathPara"))
                return true;
            return element.InnerXml.Contains("oMath");
        }
        return false;
    }

    /// <summary>
    /// Extract the OMML math element from an a14:m or mc:AlternateContent wrapper.
    /// </summary>
    private static OpenXmlElement GetMathElement(OpenXmlElement element)
    {
        if (element.LocalName == "m" && element.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main")
        {
            var child = element.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (child != null) return child;

            var desc = element.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
            if (desc != null) return desc;

            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;

            return element;
        }
        if (element is AlternateContent || element.LocalName == "AlternateContent")
        {
            var choice = element.ChildElements.FirstOrDefault(e => e is AlternateContentChoice || e.LocalName == "Choice");
            if (choice != null)
            {
                var a14m = choice.ChildElements.FirstOrDefault(e =>
                    e.LocalName == "m" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2010/main");
                if (a14m != null)
                    return GetMathElement(a14m);

                var mathDesc = choice.Descendants().FirstOrDefault(e => e.LocalName == "oMathPara" || e.LocalName == "oMath");
                if (mathDesc != null)
                    return mathDesc;
            }

            var innerXml = element.InnerXml;
            if (!string.IsNullOrEmpty(innerXml) && innerXml.Contains("oMath"))
                return ReparseFromXml(innerXml) ?? element;
        }
        return element;
    }

    /// <summary>
    /// Re-parse OMML XML string into an OpenXmlElement with navigable children.
    /// </summary>
    private static OpenXmlElement? ReparseFromXml(string innerXml)
    {
        try
        {
            var xml = innerXml.Trim();
            if (xml.Contains("oMathPara"))
            {
                var startIdx = xml.IndexOf("<m:oMathPara", StringComparison.Ordinal);
                if (startIdx < 0) startIdx = xml.IndexOf("<oMathPara", StringComparison.Ordinal);
                if (startIdx >= 0)
                {
                    var endTag = xml.Contains("</m:oMathPara>") ? "</m:oMathPara>" : "</oMathPara>";
                    var endIdx = xml.IndexOf(endTag, StringComparison.Ordinal);
                    if (endIdx >= 0)
                    {
                        var oMathParaXml = xml[startIdx..(endIdx + endTag.Length)];
                        if (!oMathParaXml.Contains("xmlns:m="))
                            oMathParaXml = oMathParaXml.Replace("<m:oMathPara", "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"");
                        var wrapper = new OpenXmlUnknownElement("m", "oMathPara", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                        var innerStart = oMathParaXml.IndexOf('>') + 1;
                        var innerEnd = oMathParaXml.LastIndexOf('<');
                        if (innerStart > 0 && innerEnd > innerStart)
                            wrapper.InnerXml = oMathParaXml[innerStart..innerEnd];
                        return wrapper;
                    }
                }
            }
        }
        catch { }
        return null;
    }

    private static bool IsTitle(Shape shape)
    {
        var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
            ?.GetFirstChild<PlaceholderShape>();
        if (ph == null) return false;
        var type = ph.Type?.Value;
        return type == PlaceholderValues.Title || type == PlaceholderValues.CenteredTitle;
    }

    private static string GetShapeName(Shape shape) =>
        shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? "?";

    private static long ParseEmu(string value) => Core.EmuConverter.ParseEmu(value);

    private static string FormatEmu(long emu) => Core.EmuConverter.FormatEmu(emu);

    private static string FormatLineWidth(long emu) => Core.EmuConverter.FormatLineWidth(emu);

    /// <summary>
    /// Normalize DrawingML alignment abbreviations to human-readable values.
    /// OOXML stores "l", "r", "ctr", "just" etc. — we return "left", "right", "center", "justify".
    /// </summary>
    private static string NormalizeAlignment(string innerText) => innerText switch
    {
        "l" => "left",
        "r" => "right",
        "ctr" => "center",
        "just" => "justify",
        "dist" => "distributed",
        _ => innerText
    };

    /// <summary>
    /// Generate a minimal 1x1 light-gray PNG for use as a zoom placeholder.
    /// PowerPoint regenerates the actual slide thumbnail when the file is opened.
    /// </summary>
    private static byte[] GenerateZoomPlaceholderPng()
    {
        // Minimal valid 1x1 PNG (RGBA: light gray #D0D0D0, fully opaque)
        using var ms = new MemoryStream();
        var bw = new BinaryWriter(ms);

        // PNG signature
        bw.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });

        // IHDR chunk: 1x1, 8-bit RGBA
        WriteChunk(bw, "IHDR", new byte[] {
            0, 0, 0, 1, // width = 1
            0, 0, 0, 1, // height = 1
            8,           // bit depth
            6,           // color type = RGBA
            0, 0, 0      // compression, filter, interlace
        });

        // IDAT chunk: zlib-compressed pixel data (filter=0, R=0xD0, G=0xD0, B=0xD0, A=0xFF)
        // Pre-computed deflate of [0x00, 0xD0, 0xD0, 0xD0, 0xFF]
        WriteChunk(bw, "IDAT", new byte[] {
            0x78, 0x01, 0x62, 0x60, 0x60, 0x28, 0x61, 0x28,
            0x61, 0x68, 0xF8, 0x0F, 0x00, 0x01, 0x45, 0x00, 0xC5
        });

        // IEND chunk
        WriteChunk(bw, "IEND", Array.Empty<byte>());

        return ms.ToArray();
    }

    private static void WriteChunk(BinaryWriter bw, string type, byte[] data)
    {
        // Length (big-endian)
        var lenBytes = BitConverter.GetBytes(data.Length);
        if (BitConverter.IsLittleEndian) Array.Reverse(lenBytes);
        bw.Write(lenBytes);

        // Type
        var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
        bw.Write(typeBytes);

        // Data
        bw.Write(data);

        // CRC32 over type + data
        var crcData = new byte[4 + data.Length];
        Array.Copy(typeBytes, 0, crcData, 0, 4);
        Array.Copy(data, 0, crcData, 4, data.Length);
        var crc = Crc32(crcData);
        var crcBytes = BitConverter.GetBytes(crc);
        if (BitConverter.IsLittleEndian) Array.Reverse(crcBytes);
        bw.Write(crcBytes);
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (var b in data)
        {
            crc ^= b;
            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (crc & 1) * 0xEDB88320;
        }
        return ~crc;
    }

    /// <summary>
    /// Find all zoom AlternateContent elements in a shape tree.
    /// </summary>
    private static List<OpenXmlElement> GetZoomElements(ShapeTree shapeTree)
    {
        return shapeTree.ChildElements
            .Where(e => e.LocalName == "AlternateContent" &&
                   e.Descendants().Any(d => d.LocalName == "sldZm"))
            .ToList();
    }

    /// <summary>
    /// Find all 3D model AlternateContent elements in a shape tree.
    /// </summary>
    private static List<OpenXmlElement> GetModel3DElements(ShapeTree shapeTree)
    {
        return shapeTree.ChildElements
            .Where(e => e.LocalName == "AlternateContent" &&
                   e.Descendants().Any(d => d.LocalName == "model3d"))
            .ToList();
    }

    /// <summary>
    /// Build a DocumentNode from a 3D model AlternateContent element.
    /// </summary>
    private DocumentNode Model3DToNode(OpenXmlElement acElement, int slideNum, int modelIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/model3d[{modelIdx}]",
            Type = "model3d"
        };

        // Navigate: mc:Choice > p:graphicFrame (or p:sp for legacy)
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var gf = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame")
              ?? choice?.ChildElements.FirstOrDefault(e => e.LocalName == "sp");

        // Name from cNvPr
        var nvGfPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr")
                  ?? gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvSpPr");
        var cNvPr = nvGfPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cNvPr != null)
        {
            var nameAttr = cNvPr.GetAttribute("name", "");
            if (!string.IsNullOrEmpty(nameAttr.Value))
                node.Format["name"] = nameAttr.Value;
        }

        // Position/size from xfrm (graphicFrame level) or spPr > xfrm
        var xfrm = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        if (xfrm == null)
        {
            var spPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "spPr");
            xfrm = spPr?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        }
        if (xfrm != null)
        {
            var off = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
            var ext = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
            if (off != null)
            {
                var xAttr = off.GetAttribute("x", "");
                var yAttr = off.GetAttribute("y", "");
                if (!string.IsNullOrEmpty(xAttr.Value) && long.TryParse(xAttr.Value, out var xVal))
                    node.Format["x"] = FormatEmu(xVal);
                if (!string.IsNullOrEmpty(yAttr.Value) && long.TryParse(yAttr.Value, out var yVal))
                    node.Format["y"] = FormatEmu(yVal);
            }
            if (ext != null)
            {
                var cxAttr = ext.GetAttribute("cx", "");
                var cyAttr = ext.GetAttribute("cy", "");
                if (!string.IsNullOrEmpty(cxAttr.Value) && long.TryParse(cxAttr.Value, out var cxVal))
                    node.Format["width"] = FormatEmu(cxVal);
                if (!string.IsNullOrEmpty(cyAttr.Value) && long.TryParse(cyAttr.Value, out var cyVal))
                    node.Format["height"] = FormatEmu(cyVal);
            }
        }

        // Model3D-specific properties
        var model3d = acElement.Descendants().FirstOrDefault(d => d.LocalName == "model3d");
        if (model3d != null)
        {
            // Model rotation
            var rot = model3d.Descendants().FirstOrDefault(d => d.LocalName == "rot");
            if (rot != null)
            {
                var ax = rot.GetAttribute("ax", "").Value ?? "";
                var ay = rot.GetAttribute("ay", "").Value ?? "";
                var az = rot.GetAttribute("az", "").Value ?? "";
                if (!string.IsNullOrEmpty(ax) || !string.IsNullOrEmpty(ay) || !string.IsNullOrEmpty(az))
                {
                    static string ToDeg(string val) =>
                        !string.IsNullOrEmpty(val) && int.TryParse(val, out var v) ? (v / 60000.0).ToString("0.##") : "0";
                    node.Format["rotation"] = $"{ToDeg(ax)},{ToDeg(ay)},{ToDeg(az)}";
                }
            }
        }

        return node;
    }

    /// <summary>
    /// Convert a SlideId value to 1-based slide number.
    /// </summary>
    private int SlideIdToNumber(uint sldId)
    {
        var slideIds = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideIdList>()
            ?.Elements<SlideId>().ToList();
        if (slideIds == null) return -1;
        for (int i = 0; i < slideIds.Count; i++)
            if (slideIds[i].Id?.Value == sldId) return i + 1;
        return -1;
    }

    /// <summary>
    /// Build a DocumentNode from a zoom AlternateContent element.
    /// </summary>
    private DocumentNode ZoomToNode(OpenXmlElement acElement, int slideNum, int zoomIdx)
    {
        var node = new DocumentNode
        {
            Path = $"/slide[{slideNum}]/zoom[{zoomIdx}]",
            Type = "zoom"
        };

        // Navigate: mc:Choice > p:graphicFrame
        var choice = acElement.ChildElements.FirstOrDefault(e => e.LocalName == "Choice");
        var gf = choice?.ChildElements.FirstOrDefault(e => e.LocalName == "graphicFrame");

        // Name from cNvPr
        var nvGfPr = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "nvGraphicFramePr");
        var cNvPr = nvGfPr?.ChildElements.FirstOrDefault(e => e.LocalName == "cNvPr");
        if (cNvPr != null)
        {
            var nameAttr = cNvPr.GetAttribute("name", "");
            if (!string.IsNullOrEmpty(nameAttr.Value))
                node.Format["name"] = nameAttr.Value;
        }

        // Position from xfrm
        var xfrm = gf?.ChildElements.FirstOrDefault(e => e.LocalName == "xfrm");
        if (xfrm != null)
        {
            var off = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "off");
            var ext = xfrm.ChildElements.FirstOrDefault(e => e.LocalName == "ext");
            if (off != null)
            {
                var xAttr = off.GetAttribute("x", "");
                var yAttr = off.GetAttribute("y", "");
                if (!string.IsNullOrEmpty(xAttr.Value) && long.TryParse(xAttr.Value, out var x))
                    node.Format["x"] = FormatEmu(x);
                if (!string.IsNullOrEmpty(yAttr.Value) && long.TryParse(yAttr.Value, out var y))
                    node.Format["y"] = FormatEmu(y);
            }
            if (ext != null)
            {
                var cxAttr = ext.GetAttribute("cx", "");
                var cyAttr = ext.GetAttribute("cy", "");
                if (!string.IsNullOrEmpty(cxAttr.Value) && long.TryParse(cxAttr.Value, out var cx))
                    node.Format["width"] = FormatEmu(cx);
                if (!string.IsNullOrEmpty(cyAttr.Value) && long.TryParse(cyAttr.Value, out var cy))
                    node.Format["height"] = FormatEmu(cy);
            }
        }

        // Zoom properties from sldZmObj / zmPr
        var sldZmObj = acElement.Descendants().FirstOrDefault(d => d.LocalName == "sldZmObj");
        if (sldZmObj != null)
        {
            var sldIdAttr = sldZmObj.GetAttribute("sldId", "");
            if (!string.IsNullOrEmpty(sldIdAttr.Value) && uint.TryParse(sldIdAttr.Value, out var sldId))
            {
                var targetNum = SlideIdToNumber(sldId);
                if (targetNum > 0) node.Format["target"] = targetNum;
            }
        }

        var zmPr = acElement.Descendants().FirstOrDefault(d => d.LocalName == "zmPr");
        if (zmPr != null)
        {
            var rtpAttr = zmPr.GetAttribute("returnToParent", "");
            if (!string.IsNullOrEmpty(rtpAttr.Value))
            {
                // Schema declares bool; normalize "1"/"0"/"true"/"false" → bool.
                node.Format["returnToParent"] = rtpAttr.Value is "1" or "true";
            }
            var tdAttr = zmPr.GetAttribute("transitionDur", "");
            if (!string.IsNullOrEmpty(tdAttr.Value))
                node.Format["transitionDur"] = tdAttr.Value;
        }

        return node;
    }

    /// <summary>
    /// Schema order for DrawingML CT_TextCharacterProperties children (a:rPr / a:endParaRPr / a:defRPr).
    /// Source: Open-XML-SDK CompositeParticle definition of TextCharacterPropertiesType.
    /// Children must appear in this order or OpenXmlValidator emits schema warnings and
    /// PowerPoint silently drops the out-of-order ones.
    /// </summary>
    private static readonly (Type type, int order)[] DrawingRunPropChildOrder = new (Type, int)[]
    {
        (typeof(Drawing.Outline),              1),   // ln
        (typeof(Drawing.NoFill),               2),   // noFill
        (typeof(Drawing.SolidFill),            2),   // solidFill
        (typeof(Drawing.GradientFill),         2),   // gradFill
        (typeof(Drawing.BlipFill),             2),   // blipFill
        (typeof(Drawing.PatternFill),          2),   // pattFill
        (typeof(Drawing.GroupFill),            2),   // grpFill
        (typeof(Drawing.EffectList),           3),   // effectLst
        (typeof(Drawing.EffectDag),            3),   // effectDag
        (typeof(Drawing.Highlight),            4),   // highlight
        (typeof(Drawing.UnderlineFollowsText), 5),   // uLnTx
        (typeof(Drawing.Underline),            5),   // uLn
        (typeof(Drawing.UnderlineFillText),    6),   // uFillTx
        (typeof(Drawing.UnderlineFill),        6),   // uFill
        (typeof(Drawing.LatinFont),            7),   // latin
        (typeof(Drawing.EastAsianFont),        8),   // ea
        (typeof(Drawing.ComplexScriptFont),    9),   // cs
        (typeof(Drawing.SymbolFont),          10),   // sym
        (typeof(Drawing.HyperlinkOnClick),    11),   // hlinkClick
        (typeof(Drawing.HyperlinkOnMouseOver),12),   // hlinkMouseOver
        (typeof(Drawing.RightToLeft),         13),   // rtl
        (typeof(Drawing.ExtensionList),       14),   // extLst
    };

    /// <summary>
    /// Reorder children of a DrawingML RunProperties / EndParagraphRunProperties /
    /// DefaultRunProperties element into schema-valid order.
    /// Stable within the same order bucket to preserve relative order of existing fills.
    /// Unknown child types are pushed to the end (preserved but last).
    /// </summary>
    internal static void ReorderDrawingRunProperties(OpenXmlCompositeElement rPr)
    {
        if (rPr == null || !rPr.HasChildren) return;

        int OrderOf(OpenXmlElement el)
        {
            var t = el.GetType();
            foreach (var (type, order) in DrawingRunPropChildOrder)
                if (type == t) return order;
            return int.MaxValue;
        }

        var children = rPr.ChildElements.ToList();
        // Check if already sorted — avoid unnecessary reflows
        bool needsReorder = false;
        for (int i = 1; i < children.Count; i++)
        {
            if (OrderOf(children[i]) < OrderOf(children[i - 1]))
            {
                needsReorder = true;
                break;
            }
        }
        if (!needsReorder) return;

        // Stable sort by schema order
        var sorted = children
            .Select((el, idx) => (el, ord: OrderOf(el), idx))
            .OrderBy(t => t.ord)
            .ThenBy(t => t.idx)
            .Select(t => t.el)
            .ToList();

        foreach (var c in children) c.Remove();
        foreach (var c in sorted) rPr.AppendChild(c);
    }

    /// <summary>
    /// Read a GradientFill element and return a string representation (C1-C2[-angle] or radial:C1-C2[-focus]).
    /// </summary>
    /// <summary>
    /// Read a gradient stop color, handling both RgbColorModelHex and SchemeColor.
    /// Without this, scheme-color stops (accent1/dark1/...) read back as "#?" because
    /// FormatHexColor receives the literal "?" placeholder.
    /// </summary>
    private static string ReadGradientStopColor(Drawing.GradientStop gs)
    {
        var rgb = gs.GetFirstChild<Drawing.RgbColorModelHex>();
        if (rgb?.Val?.Value != null) return ParseHelpers.FormatHexColor(rgb.Val.Value);
        var scheme = gs.GetFirstChild<Drawing.SchemeColor>();
        if (scheme?.Val?.Value != null) return scheme.Val.Value.ToString();
        var sys = gs.GetFirstChild<Drawing.SystemColor>();
        if (sys?.Val?.Value != null) return sys.Val.Value.ToString();
        var preset = gs.GetFirstChild<Drawing.PresetColor>();
        if (preset?.Val?.Value != null) return preset.Val.Value.ToString();
        return "?";
    }

    internal static string ReadGradientString(Drawing.GradientFill gradFill)
    {
        var stopEls = gradFill.GradientStopList?.Elements<Drawing.GradientStop>().ToList();
        if (stopEls == null || stopEls.Count == 0) return "gradient";

        var stopData = stopEls.Select(gs => (
            color: ReadGradientStopColor(gs),
            pos: gs.Position?.Value
        )).ToList();

        // Check if positions deviate >1% from even distribution (1000 units)
        bool hasCustomPos = false;
        int n = stopData.Count;
        for (int i = 0; i < n; i++)
        {
            var expectedPos = n == 1 ? 0 : (int)((long)i * 100000 / (n - 1));
            var actualPos = (int)(stopData[i].pos ?? 0);
            if (Math.Abs(actualPos - expectedPos) > 1000) { hasCustomPos = true; break; }
        }

        var stopStrs = stopData.Select((s, i) =>
            hasCustomPos && s.pos.HasValue
                ? $"{s.color}@{s.pos.Value / 1000}"
                : s.color
        ).ToList();

        var pathGrad = gradFill.GetFirstChild<Drawing.PathGradientFill>();
        if (pathGrad != null)
        {
            var fillRect = pathGrad.GetFirstChild<Drawing.FillToRectangle>();
            var focus = "center";
            if (fillRect != null)
            {
                var fl = fillRect.Left?.Value ?? 50000;
                var ft = fillRect.Top?.Value ?? 50000;
                focus = (fl, ft) switch
                {
                    (0, 0) => "tl",
                    ( >= 100000, 0) => "tr",
                    (0, >= 100000) => "bl",
                    ( >= 100000, >= 100000) => "br",
                    _ => "center"
                };
            }
            return $"radial:{string.Join("-", stopStrs)}-{focus}";
        }

        var linear = gradFill.GetFirstChild<Drawing.LinearGradientFill>();
        var deg = linear?.Angle?.HasValue == true ? linear.Angle.Value / 60000.0 : 0.0;
        var degStr = deg % 1 == 0 ? $"{(int)deg}" : $"{deg:0.##}";
        return $"linear;{string.Join(";", stopStrs)};{degStr}";
    }

    /// <summary>
    /// Parse SVG-like path syntax into a Drawing.CustomGeometry element.
    /// Format: "M x,y L x,y C x1,y1 x2,y2 x,y Q x1,y1 x,y Z"
    ///   M = moveTo, L = lineTo, C = cubicBezTo, Q = quadBezTo, A = arcTo, Z = close
    /// Coordinates use 0-100 relative space, internally scaled ×1000 to OOXML standard 0-100000.
    /// Example: "M 0,0 L 100,0 L 100,100 L 0,100 Z" (rectangle in 0-100 space)
    /// </summary>
    private static Drawing.CustomGeometry ParseCustomGeometry(string value)
    {
        var path = new Drawing.Path();

        // Parse SVG-like commands
        var tokens = value.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        long maxX = 0, maxY = 0;
        int i = 0;

        while (i < tokens.Length)
        {
            var cmd = tokens[i].ToUpperInvariant();
            i++;

            switch (cmd)
            {
                case "M":
                {
                    var (x, y) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.MoveTo(new Drawing.Point { X = x.ToString(), Y = y.ToString() }));
                    TrackMax(ref maxX, ref maxY, x, y);
                    break;
                }
                case "L":
                {
                    var (x, y) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.LineTo(new Drawing.Point { X = x.ToString(), Y = y.ToString() }));
                    TrackMax(ref maxX, ref maxY, x, y);
                    break;
                }
                case "C":
                {
                    // Cubic bezier: 3 points (control1, control2, end)
                    var (x1, y1) = ParsePointToken(tokens[i++]);
                    var (x2, y2) = ParsePointToken(tokens[i++]);
                    var (x3, y3) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.CubicBezierCurveTo(
                        new Drawing.Point { X = x1.ToString(), Y = y1.ToString() },
                        new Drawing.Point { X = x2.ToString(), Y = y2.ToString() },
                        new Drawing.Point { X = x3.ToString(), Y = y3.ToString() }
                    ));
                    TrackMax(ref maxX, ref maxY, x3, y3);
                    break;
                }
                case "Q":
                {
                    // Quadratic bezier: 2 points (control, end)
                    var (x1, y1) = ParsePointToken(tokens[i++]);
                    var (x2, y2) = ParsePointToken(tokens[i++]);
                    path.AppendChild(new Drawing.QuadraticBezierCurveTo(
                        new Drawing.Point { X = x1.ToString(), Y = y1.ToString() },
                        new Drawing.Point { X = x2.ToString(), Y = y2.ToString() }
                    ));
                    TrackMax(ref maxX, ref maxY, x2, y2);
                    break;
                }
                case "Z":
                    path.AppendChild(new Drawing.CloseShapePath());
                    break;
                default:
                    // Skip unknown tokens
                    break;
            }
        }

        // Set path dimensions to bounding box
        if (maxX > 0) path.Width = maxX;
        if (maxY > 0) path.Height = maxY;

        return new Drawing.CustomGeometry(
            new Drawing.AdjustValueList(),
            new Drawing.ShapeGuideList(),
            new Drawing.AdjustHandleList(),
            new Drawing.ConnectionSiteList(),
            new Drawing.Rectangle { Left = "0", Top = "0", Right = "r", Bottom = "b" },
            new Drawing.PathList(path)
        );
    }

    /// <summary>
    /// Parse "x,y" coordinate token and scale ×1000 to OOXML standard 0-100000 range.
    /// Input coordinates are 0-100 relative space.
    /// </summary>
    private static (long x, long y) ParsePointToken(string token)
    {
        var parts = token.Split(',');
        if (parts.Length < 2)
            throw new ArgumentException($"Invalid coordinate '{token}'. Expected 'x,y' format (e.g. '100,200').");
        if (!long.TryParse(parts[0].Trim(), out var x))
            throw new ArgumentException($"Invalid x coordinate '{parts[0].Trim()}' in '{token}'. Expected a number.");
        if (!long.TryParse(parts[1].Trim(), out var y))
            throw new ArgumentException($"Invalid y coordinate '{parts[1].Trim()}' in '{token}'. Expected a number.");
        // Scale from user space (0-100) to OOXML standard (0-100000)
        return (x * 1000, y * 1000);
    }

    private static void TrackMax(ref long maxX, ref long maxY, long x, long y)
    {
        if (x > maxX) maxX = x;
        if (y > maxY) maxY = y;
    }

    /// <summary>
    /// Change the z-order of a shape within the ShapeTree.
    /// Values: "front" (topmost), "back" (bottommost), "forward" (+1), "backward" (-1),
    ///         or an integer for absolute position (1-based, 1 = back, N = front).
    /// </summary>
    private static void ApplyZOrder(DocumentFormat.OpenXml.Packaging.SlidePart slidePart, Shape shape, string value)
    {
        var shapeTree = shape.Parent as ShapeTree
            ?? throw new InvalidOperationException("Shape is not in a ShapeTree");

        // Get all content elements (Shape, Picture, GraphicFrame, GroupShape, ConnectionShape)
        // that participate in z-order (skip structural elements like nvGrpSpPr, grpSpPr)
        var contentElements = shapeTree.ChildElements
            .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
            .ToList();
        var currentIndex = contentElements.IndexOf(shape);
        if (currentIndex < 0) return;

        int targetIndex;
        switch (value.ToLowerInvariant())
        {
            case "front" or "top" or "bringtofront":
                targetIndex = contentElements.Count - 1;
                break;
            case "back" or "bottom" or "sendtoback":
                targetIndex = 0;
                break;
            case "forward" or "bringforward" or "+1":
                targetIndex = Math.Min(currentIndex + 1, contentElements.Count - 1);
                break;
            case "backward" or "sendbackward" or "-1":
                targetIndex = Math.Max(currentIndex - 1, 0);
                break;
            default:
                // Absolute position (1-based: 1 = back, N = front)
                if (int.TryParse(value, out var pos))
                    targetIndex = Math.Clamp(pos - 1, 0, contentElements.Count - 1);
                else
                    throw new ArgumentException($"Invalid z-order value: {value}. Use front/back/forward/backward or a number.");
                break;
        }

        if (targetIndex == currentIndex) return;

        // Remove shape from its current position
        shape.Remove();

        // Insert at new position
        if (targetIndex >= contentElements.Count - 1)
        {
            // Front: append after last content element (or at end of tree)
            shapeTree.AppendChild(shape);
        }
        else if (targetIndex <= 0)
        {
            // Back: insert before the first content element
            var firstContent = shapeTree.ChildElements
                .FirstOrDefault(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape);
            if (firstContent != null)
                firstContent.InsertBeforeSelf(shape);
            else
                shapeTree.AppendChild(shape);
        }
        else
        {
            // Refresh content list after removal
            var updatedContent = shapeTree.ChildElements
                .Where(e => e is Shape or Picture or GraphicFrame or GroupShape or ConnectionShape)
                .ToList();
            if (targetIndex < updatedContent.Count)
                updatedContent[targetIndex].InsertBeforeSelf(shape);
            else
                shapeTree.AppendChild(shape);
        }
    }

    /// <summary>
    /// Apply a position/size property (x, y, width, height) to offset and extents.
    /// Returns true if the key was handled.
    /// </summary>
    private static bool TryApplyPositionSize(string key, string value, Drawing.Offset offset, Drawing.Extents extents)
    {
        var emu = ParseEmu(value);
        switch (key)
        {
            case "x": offset.X = emu; return true;
            case "y": offset.Y = emu; return true;
            case "width":
                if (emu < 0) throw new ArgumentException($"Negative width is not allowed: '{value}'.");
                extents.Cx = emu; return true;
            case "height":
                if (emu < 0) throw new ArgumentException($"Negative height is not allowed: '{value}'.");
                extents.Cy = emu; return true;
            default: return false;
        }
    }

    private static readonly Dictionary<string, string> _tableStyleNameToGuid = new(StringComparer.OrdinalIgnoreCase)
    {
        ["medium1"] = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
        ["mediumstyle1"] = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
        ["medium2"] = "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
        ["mediumstyle2"] = "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
        ["medium3"] = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
        ["mediumstyle3"] = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
        ["medium4"] = "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
        ["mediumstyle4"] = "{D7AC3CCA-C797-4891-BE02-D94E43425B78}",
        ["light1"] = "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
        ["lightstyle1"] = "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}",
        ["light2"] = "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
        ["lightstyle2"] = "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}",
        ["light3"] = "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
        ["lightstyle3"] = "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}",
        ["dark1"] = "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
        ["darkstyle1"] = "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}",
        ["dark2"] = "{125E5076-3810-47DD-B79F-674D7AD40C01}",
        ["darkstyle2"] = "{125E5076-3810-47DD-B79F-674D7AD40C01}",
        ["none"] = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
    };

    /// <summary>
    /// Resolve a table style name or GUID to a valid OOXML GUID.
    /// Throws ArgumentException for unrecognized style names.
    /// </summary>
    private static string ResolveTableStyleId(string value)
    {
        if (_tableStyleNameToGuid.TryGetValue(value, out var guid))
            return guid;
        if (value.StartsWith("{"))
            return value; // Direct GUID passthrough
        throw new ArgumentException(
            $"Invalid table style: '{value}'. Valid values: medium1, medium2, medium3, medium4, light1, light2, light3, dark1, dark2, none, or a direct GUID like {{073A0DAA-...}}.");
    }

    /// <summary>
    /// Find and replace text across all slides. Returns the number of replacements made.
    /// </summary>
    // ==================== Find / Format / Replace ====================

    /// <summary>
    /// Build a flat list of (Run, Text, charStart, charEnd) spans for a PPT paragraph.
    /// </summary>
    private static List<(Drawing.Run Run, Drawing.Text TextElement, int Start, int End)> BuildPptRunTexts(Drawing.Paragraph para)
    {
        var runTexts = new List<(Drawing.Run Run, Drawing.Text TextElement, int Start, int End)>();
        int pos = 0;
        foreach (var run in para.Descendants<Drawing.Run>())
        {
            var text = run.GetFirstChild<Drawing.Text>();
            var len = text?.Text?.Length ?? 0;
            if (len > 0)
                runTexts.Add((run, text!, pos, pos + len));
            pos += len;
        }
        return runTexts;
    }

    /// <summary>
    /// Parse a find pattern: plain text or regex (r"..." prefix).
    /// </summary>
    private static (string Pattern, bool IsRegex) ParseFindPattern(string value)
    {
        if (value.Length >= 3 && value[0] == 'r' && (value[1] == '"' || value[1] == '\''))
        {
            var quote = value[1];
            var endIdx = value.LastIndexOf(quote);
            if (endIdx > 1)
                return (value[2..endIdx], true);
        }
        return (value, false);
    }

    /// <summary>
    /// Find all match ranges in fullText using either plain text or regex.
    /// </summary>
    private static List<(int Start, int Length)> FindMatchRanges(string fullText, string pattern, bool isRegex)
    {
        var ranges = new List<(int Start, int Length)>();
        if (isRegex)
        {
            try
            {
                // BUG-TESTER fuzz-2: bound matching with hard timeout to prevent
                // catastrophic-backtracking DoS.
                foreach (Match m in Regex.Matches(fullText, pattern, RegexOptions.None, FindRegexMatchTimeout))
                {
                    if (m.Length > 0)
                        ranges.Add((m.Index, m.Length));
                }
            }
            catch (RegexParseException ex)
            {
                throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}", ex);
            }
            catch (RegexMatchTimeoutException ex)
            {
                throw new ArgumentException(
                    $"Regex pattern '{pattern}' exceeded {FindRegexMatchTimeout.TotalSeconds}s match timeout (catastrophic backtracking?)",
                    ex);
            }
        }
        else
        {
            int idx = 0;
            while ((idx = fullText.IndexOf(pattern, idx, StringComparison.Ordinal)) >= 0)
            {
                ranges.Add((idx, pattern.Length));
                idx += pattern.Length;
            }
        }
        return ranges;
    }

    /// <summary>
    /// Split a PPT run at a character offset. Returns the new right-side run.
    /// RunProperties are deep-cloned.
    /// </summary>
    private static Drawing.Run SplitPptRunAtOffset(Drawing.Run run, int charOffset)
    {
        var text = run.GetFirstChild<Drawing.Text>();
        if (text?.Text == null || charOffset <= 0 || charOffset >= text.Text.Length)
            return run;

        var leftText = text.Text[..charOffset];
        var rightText = text.Text[charOffset..];

        // Clone the run for the right side
        var rightRun = (Drawing.Run)run.CloneNode(true);

        // Set text
        text.Text = leftText;
        var rightTextElem = rightRun.GetFirstChild<Drawing.Text>();
        if (rightTextElem != null) rightTextElem.Text = rightText;

        // Insert after original
        run.InsertAfterSelf(rightRun);
        return rightRun;
    }

    /// <summary>
    /// Split runs in a PPT paragraph so that [charStart, charEnd) is covered by dedicated runs.
    /// Returns the runs covering that range.
    /// </summary>
    private static List<Drawing.Run> SplitPptRunsAtRange(Drawing.Paragraph para, int charStart, int charEnd)
    {
        // Split at charEnd first
        var runTexts = BuildPptRunTexts(para);
        foreach (var rt in runTexts)
        {
            if (charEnd > rt.Start && charEnd < rt.End)
            {
                SplitPptRunAtOffset(rt.Run, charEnd - rt.Start);
                break;
            }
        }

        // Rebuild, then split at charStart
        runTexts = BuildPptRunTexts(para);
        foreach (var rt in runTexts)
        {
            if (charStart > rt.Start && charStart < rt.End)
            {
                SplitPptRunAtOffset(rt.Run, charStart - rt.Start);
                break;
            }
        }

        // Collect runs covering [charStart, charEnd)
        runTexts = BuildPptRunTexts(para);
        var result = new List<Drawing.Run>();
        foreach (var rt in runTexts)
        {
            if (rt.Start >= charStart && rt.End <= charEnd)
                result.Add(rt.Run);
        }
        return result;
    }

    /// <summary>
    /// Apply run-level formatting to a PPT run's RunProperties.
    /// </summary>
    private static void ApplyPptRunFormatting(Drawing.Run run, string key, string value, Shape? shape = null)
    {
        var rPr = run.RunProperties ?? run.PrependChild(new Drawing.RunProperties());
        switch (key.ToLowerInvariant())
        {
            case "bold":
                rPr.Bold = IsTruthy(value);
                break;
            case "italic":
                rPr.Italic = IsTruthy(value);
                break;
            case "size":
                rPr.FontSize = (int)Math.Round(ParseFontSize(value) * 100, MidpointRounding.AwayFromZero);
                break;
            case "color":
                rPr.RemoveAllChildren<Drawing.SolidFill>();
                rPr.PrependChild(BuildSolidFill(value));
                break;
            case "font":
                rPr.RemoveAllChildren<Drawing.LatinFont>();
                rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                break;
            case "underline":
                var ulVal = value.ToLowerInvariant() switch
                {
                    "true" or "single" => Drawing.TextUnderlineValues.Single,
                    "double" => Drawing.TextUnderlineValues.Double,
                    "heavy" => Drawing.TextUnderlineValues.Heavy,
                    "false" or "none" => Drawing.TextUnderlineValues.None,
                    _ => new Drawing.TextUnderlineValues(value)
                };
                rPr.Underline = ulVal;
                break;
            case "strikethrough" or "strike":
                var stVal = value.ToLowerInvariant() switch
                {
                    "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                    "double" => Drawing.TextStrikeValues.DoubleStrike,
                    "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                    _ => new Drawing.TextStrikeValues(value)
                };
                rPr.Strike = stVal;
                break;
            case "superscript":
                rPr.Baseline = IsTruthy(value) ? 30000 : 0;
                break;
            case "subscript":
                rPr.Baseline = IsTruthy(value) ? -25000 : 0;
                break;
            case "charspacing" or "spacing" or "letterspacing":
                var csPt = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? ParseHelpers.SafeParseDouble(value[..^2], "charspacing")
                    : ParseHelpers.SafeParseDouble(value, "charspacing");
                rPr.Spacing = (int)Math.Round(csPt * 100, MidpointRounding.AwayFromZero);
                break;
            case "highlight":
                rPr.RemoveAllChildren<Drawing.Highlight>();
                if (!string.Equals(value, "none", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
                {
                    var hl = new Drawing.Highlight();
                    hl.AppendChild(BuildSolidFillColor(value));
                    rPr.AppendChild(hl);
                }
                break;
        }
    }

    /// <summary>
    /// Process find in a single PPT paragraph: replace text and/or apply formatting.
    /// </summary>
    private static int ProcessFindInPptParagraph(
        Drawing.Paragraph para,
        string pattern,
        bool isRegex,
        string? replace,
        Dictionary<string, string>? formatProps,
        Shape? shape = null,
        int? runIndexFilter = null)
    {
        var runTexts = BuildPptRunTexts(para);
        if (runTexts.Count == 0) return 0;

        // BUG-TESTER+FUZZER R32: when scope is /r[K], restrict find to that
        // run's text range only. Out-of-bound was already rejected upstream.
        int scanStart = 0;
        int scanEnd = runTexts[^1].End;
        if (runIndexFilter.HasValue)
        {
            if (runIndexFilter.Value < 1 || runIndexFilter.Value > runTexts.Count)
                return 0;
            scanStart = runTexts[runIndexFilter.Value - 1].Start;
            scanEnd = runTexts[runIndexFilter.Value - 1].End;
        }

        var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
        // CONSISTENCY(regex-backref-expand): mirror Word ProcessFindInParagraph.
        // BUG-TESTER+FUZZER R31: wrap with try/catch so RegexMatchTimeoutException is
        // converted to ArgumentException, and avoid a second Regex.Matches call by
        // deriving ranges from the same Match list.
        List<System.Text.RegularExpressions.Match>? matchObjs = null;
        List<(int Start, int Length)> matches;
        if (isRegex)
        {
            try
            {
                matchObjs = System.Text.RegularExpressions.Regex.Matches(
                        fullText,
                        pattern,
                        System.Text.RegularExpressions.RegexOptions.None,
                        FindRegexMatchTimeout)
                    .Cast<System.Text.RegularExpressions.Match>()
                    .Where(m => m.Length > 0)
                    .ToList();
            }
            catch (System.Text.RegularExpressions.RegexParseException ex)
            {
                throw new ArgumentException($"Invalid regex pattern '{pattern}': {ex.Message}", ex);
            }
            catch (System.Text.RegularExpressions.RegexMatchTimeoutException ex)
            {
                throw new ArgumentException(
                    $"Regex pattern '{pattern}' exceeded {FindRegexMatchTimeout.TotalSeconds}s match timeout (catastrophic backtracking?)",
                    ex);
            }
            matches = matchObjs.Select(m => (m.Index, m.Length)).ToList();
        }
        else
        {
            matches = FindMatchRanges(fullText, pattern, isRegex);
        }

        // Apply run-scope filter (R32): keep only matches fully contained in the run.
        if (runIndexFilter.HasValue)
        {
            var keepIdx = new HashSet<int>();
            for (int k = 0; k < matches.Count; k++)
            {
                var (s, l) = matches[k];
                if (s >= scanStart && s + l <= scanEnd)
                    keepIdx.Add(k);
            }
            matches = matches.Where((_, k) => keepIdx.Contains(k)).ToList();
            if (matchObjs != null)
                matchObjs = matchObjs.Where((_, k) => keepIdx.Contains(k)).ToList();
        }

        if (matches.Count == 0) return 0;

        for (int i = matches.Count - 1; i >= 0; i--)
        {
            var (matchStart, matchLen) = matches[i];
            var matchEnd = matchStart + matchLen;

            if (replace != null)
            {
                // Expand backrefs via Match.Result so lookarounds keep their context.
                string effectiveReplace = replace;
                if (isRegex && matchObjs != null && i < matchObjs.Count)
                {
                    effectiveReplace = matchObjs[i].Result(replace);
                }

                // Replace text in affected runs
                var currentRunTexts = BuildPptRunTexts(para);
                bool first = true;
                foreach (var rt in currentRunTexts)
                {
                    if (rt.End <= matchStart || rt.Start >= matchEnd)
                        continue;

                    var textStr = rt.TextElement.Text ?? "";
                    var localStart = Math.Max(0, matchStart - rt.Start);
                    var localEnd = Math.Min(textStr.Length, matchEnd - rt.Start);

                    if (first)
                    {
                        rt.TextElement.Text = textStr[..localStart] + effectiveReplace + textStr[localEnd..];
                        first = false;
                    }
                    else
                    {
                        rt.TextElement.Text = textStr[..Math.Max(0, matchStart - rt.Start)] + textStr[localEnd..];
                    }
                }

                // BUG-TESTER fuzz-1 (PPTX mirror): drop orphan empty <a:r> runs left
                // by cross-run replace. Only remove runs with empty <a:t> and no other
                // semantic children (RunProperties alone is not semantic content).
                var emptyRunsToRemove = new List<Drawing.Run>();
                foreach (var run in para.Descendants<Drawing.Run>())
                {
                    bool hasContent = false;
                    bool hasEmptyText = false;
                    foreach (var child in run.ChildElements)
                    {
                        if (child is Drawing.RunProperties)
                            continue;
                        if (child is Drawing.Text t)
                        {
                            if (string.IsNullOrEmpty(t.Text))
                                hasEmptyText = true;
                            else
                                hasContent = true;
                        }
                        else
                        {
                            hasContent = true;
                        }
                    }
                    if (hasEmptyText && !hasContent)
                        emptyRunsToRemove.Add(run);
                }
                foreach (var run in emptyRunsToRemove)
                    run.Remove();

                if (formatProps != null && formatProps.Count > 0 && effectiveReplace.Length > 0)
                {
                    var replacedEnd = matchStart + effectiveReplace.Length;
                    var targetRuns = SplitPptRunsAtRange(para, matchStart, replacedEnd);
                    foreach (var run in targetRuns)
                        foreach (var (key, value) in formatProps)
                            ApplyPptRunFormatting(run, key, value, shape);
                }
            }
            else if (formatProps != null && formatProps.Count > 0)
            {
                var targetRuns = SplitPptRunsAtRange(para, matchStart, matchEnd);
                foreach (var run in targetRuns)
                    foreach (var (key, value) in formatProps)
                        ApplyPptRunFormatting(run, key, value, shape);
            }
        }

        return matches.Count;
    }

    /// <summary>
    /// Unified find across all paragraphs in the resolved scope.
    /// </summary>
    private int ProcessPptFind(string path, string findValue, string? replace, Dictionary<string, string> formatProps)
    {
        var (pattern, isRegex) = ParseFindPattern(findValue);
        if (string.IsNullOrEmpty(pattern) && !isRegex) return 0;

        int totalCount = 0;

        if (path is "/" or "" or "/presentation")
        {
            // All slides
            foreach (var slidePart in _doc.PresentationPart?.SlideParts ?? Enumerable.Empty<SlidePart>())
            {
                var slide = slidePart.Slide;
                if (slide == null) continue;
                foreach (var para in slide.Descendants<Drawing.Paragraph>())
                    totalCount += ProcessFindInPptParagraph(para, pattern, isRegex, replace,
                        formatProps.Count > 0 ? formatProps : null);
                slidePart.Slide!.Save();
            }
        }
        else
        {
            // Path-scoped: resolve to specific paragraphs (and optional run filter)
            var (paragraphs, runIndex) = ResolvePptParagraphsForFindInternal(path);
            Shape? contextShape = null;
            // Try to resolve shape for color context (anchored shape segment only).
            var shapeMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\](?:/|$)");
            if (shapeMatch.Success && shapeMatch.Groups[2].Value is not ("table" or "notes"))
            {
                try
                {
                    var (_, shape) = ResolveShape(int.Parse(shapeMatch.Groups[1].Value), int.Parse(shapeMatch.Groups[3].Value));
                    contextShape = shape;
                }
                catch { }
            }

            foreach (var para in paragraphs)
                totalCount += ProcessFindInPptParagraph(para, pattern, isRegex, replace,
                    formatProps.Count > 0 ? formatProps : null, contextShape, runIndex);

            // Save affected slides
            foreach (var slidePart in _doc.PresentationPart?.SlideParts ?? Enumerable.Empty<SlidePart>())
                slidePart.Slide?.Save();
        }

        return totalCount;
    }

    /// <summary>
    /// Resolve paragraphs from a PPT path for find operations.
    /// BUG-TESTER+FUZZER R32: paths must match exactly (anchored). Out-of-bound
    /// indices and unrecognized PPT paths throw ArgumentException instead of
    /// silently falling back to a wider scope (e.g. all slides).
    /// </summary>
    private List<Drawing.Paragraph> ResolvePptParagraphsForFind(string path)
    {
        var (paragraphs, _) = ResolvePptParagraphsForFindInternal(path);
        return paragraphs;
    }

    /// <summary>
    /// Resolve paragraphs and an optional 1-based run filter from a PPT path.
    /// When the path ends with /r[R] or /run[R], only that run within the
    /// resolved paragraph participates in find/replace.
    /// </summary>
    private (List<Drawing.Paragraph> Paragraphs, int? RunIndex) ResolvePptParagraphsForFindInternal(string path)
    {
        var paragraphs = new List<Drawing.Paragraph>();

        // /slide[N]/notes → paragraphs in notes slide
        var notesMatch = Regex.Match(path, @"^/slide\[(\d+)\]/notes$", RegexOptions.IgnoreCase);
        if (notesMatch.Success)
        {
            var slideIdx = int.Parse(notesMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide index out of range: {slideIdx} (have {slideParts.Count} slides)");
            var notesPart = slideParts[slideIdx - 1].NotesSlidePart;
            if (notesPart?.NotesSlide != null)
                paragraphs.AddRange(notesPart.NotesSlide.Descendants<Drawing.Paragraph>());
            return (paragraphs, null);
        }

        // /slide[N]/table[M]/tr[R]/tc[C][/p[P][/r[K]]] → paragraphs in table cell
        var tableCellMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]/tr\[(\d+)\]/tc\[(\d+)\](?:/p(?:aragraph)?\[(\d+)\](?:/r(?:un)?\[(\d+)\])?)?$");
        if (tableCellMatch.Success)
        {
            var slideIdx = int.Parse(tableCellMatch.Groups[1].Value);
            var tableIdx = int.Parse(tableCellMatch.Groups[2].Value);
            var rowIdx = int.Parse(tableCellMatch.Groups[3].Value);
            var colIdx = int.Parse(tableCellMatch.Groups[4].Value);
            int? paraIdx = tableCellMatch.Groups[5].Success ? int.Parse(tableCellMatch.Groups[5].Value) : (int?)null;
            int? runIdx = tableCellMatch.Groups[6].Success ? int.Parse(tableCellMatch.Groups[6].Value) : (int?)null;
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide index out of range: {slideIdx}");
            var slide = slideParts[slideIdx - 1].Slide;
            var tables = slide?.Descendants<Drawing.Table>().ToList() ?? new List<Drawing.Table>();
            if (tableIdx < 1 || tableIdx > tables.Count)
                throw new ArgumentException($"Table index out of range: {tableIdx}");
            var rows = tables[tableIdx - 1].Elements<Drawing.TableRow>().ToList();
            if (rowIdx < 1 || rowIdx > rows.Count)
                throw new ArgumentException($"Row index out of range: {rowIdx}");
            var cells = rows[rowIdx - 1].Elements<Drawing.TableCell>().ToList();
            if (colIdx < 1 || colIdx > cells.Count)
                throw new ArgumentException($"Column index out of range: {colIdx}");
            var cellParas = cells[colIdx - 1].Descendants<Drawing.Paragraph>().ToList();
            if (paraIdx.HasValue)
            {
                if (paraIdx.Value < 1 || paraIdx.Value > cellParas.Count)
                    throw new ArgumentException($"Paragraph index out of range: {paraIdx.Value} (cell has {cellParas.Count})");
                paragraphs.Add(cellParas[paraIdx.Value - 1]);
            }
            else
            {
                paragraphs.AddRange(cellParas);
            }
            if (runIdx.HasValue)
            {
                var runCount = paragraphs[0].Descendants<Drawing.Run>().Count(r => (r.GetFirstChild<Drawing.Text>()?.Text?.Length ?? 0) > 0);
                if (runIdx.Value < 1 || runIdx.Value > runCount)
                    throw new ArgumentException($"Run index out of range: {runIdx.Value} (paragraph has {runCount} runs)");
            }
            return (paragraphs, runIdx);
        }

        // /slide[N]/table[M] → all paragraphs in table
        var tableMatch = Regex.Match(path, @"^/slide\[(\d+)\]/table\[(\d+)\]$");
        if (tableMatch.Success)
        {
            var slideIdx = int.Parse(tableMatch.Groups[1].Value);
            var tableIdx = int.Parse(tableMatch.Groups[2].Value);
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide index out of range: {slideIdx}");
            var slide = slideParts[slideIdx - 1].Slide;
            var tables = slide?.Descendants<Drawing.Table>().ToList() ?? new List<Drawing.Table>();
            if (tableIdx < 1 || tableIdx > tables.Count)
                throw new ArgumentException($"Table index out of range: {tableIdx}");
            paragraphs.AddRange(tables[tableIdx - 1].Descendants<Drawing.Paragraph>());
            return (paragraphs, null);
        }

        // /slide[N]/<shape>[M][/p[P][/r[K]]] — shape with optional paragraph/run suffix
        // BUG-TESTER+FUZZER R32: anchored ($) so /p[P] suffix is not silently
        // swallowed as a prefix match against the shape selector.
        var shapeMatch = Regex.Match(path, @"^/slide\[(\d+)\]/(\w+)\[(\d+)\](?:/p(?:aragraph)?\[(\d+)\](?:/r(?:un)?\[(\d+)\])?)?$");
        if (shapeMatch.Success)
        {
            var slideIdx = int.Parse(shapeMatch.Groups[1].Value);
            var shapeKind = shapeMatch.Groups[2].Value;
            // Reject path segments that are not shape-like containers handled here.
            if (shapeKind is "table" or "notes")
                throw new ArgumentException($"Unsupported find scope path: {path}");
            var shapeIdx = int.Parse(shapeMatch.Groups[3].Value);
            int? paraIdx = shapeMatch.Groups[4].Success ? int.Parse(shapeMatch.Groups[4].Value) : (int?)null;
            int? runIdx = shapeMatch.Groups[5].Success ? int.Parse(shapeMatch.Groups[5].Value) : (int?)null;
            Shape shape;
            try
            {
                (_, shape) = ResolveShape(slideIdx, shapeIdx);
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Cannot resolve shape at {path}: {ex.Message}", ex);
            }
            if (shape.TextBody == null)
                return (paragraphs, null);
            var shapeParas = shape.TextBody.Elements<Drawing.Paragraph>().ToList();
            if (paraIdx.HasValue)
            {
                if (paraIdx.Value < 1 || paraIdx.Value > shapeParas.Count)
                    throw new ArgumentException($"Paragraph index out of range: {paraIdx.Value} (shape has {shapeParas.Count})");
                paragraphs.Add(shapeParas[paraIdx.Value - 1]);
            }
            else
            {
                paragraphs.AddRange(shapeParas);
            }
            if (runIdx.HasValue)
            {
                var runCount = paragraphs[0].Descendants<Drawing.Run>().Count(r => (r.GetFirstChild<Drawing.Text>()?.Text?.Length ?? 0) > 0);
                if (runIdx.Value < 1 || runIdx.Value > runCount)
                    throw new ArgumentException($"Run index out of range: {runIdx.Value} (paragraph has {runCount} runs)");
            }
            return (paragraphs, runIdx);
        }

        // /slide[N] → all paragraphs in slide
        var slideOnlyMatch = Regex.Match(path, @"^/slide\[(\d+)\]$");
        if (slideOnlyMatch.Success)
        {
            var slideIdx = int.Parse(slideOnlyMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (slideIdx < 1 || slideIdx > slideParts.Count)
                throw new ArgumentException($"Slide index out of range: {slideIdx}");
            var slide = slideParts[slideIdx - 1].Slide;
            if (slide != null)
                paragraphs.AddRange(slide.Descendants<Drawing.Paragraph>());
            return (paragraphs, null);
        }

        // BUG-FUZZER R32: unrecognized PPT path (e.g. /body) must not silently
        // fall back to all-slides global scope. Reject it.
        throw new ArgumentException($"Unrecognized PPT find scope path: '{path}'. Expected /, /slide[N], /slide[N]/<shape>[M][/p[P][/r[K]]], /slide[N]/notes, or /slide[N]/table[M][/tr[R]/tc[C]].");
    }

    /// <summary>
    /// Build a color element for PPT highlight from a color value.
    /// </summary>
    private static Drawing.RgbColorModelHex BuildSolidFillColor(string value)
    {
        var hex = ParseHelpers.NormalizeArgbColor(value);
        return new Drawing.RgbColorModelHex { Val = hex };
    }

    /// <summary>
    /// Add an element at a text-find position within a PPT paragraph.
    /// For PPT, this only supports inline types (run) — splits the run at the find position.
    /// </summary>
    private string AddPptAtFindPosition(
        string parentPath,
        string type,
        string findValue,
        bool isAfter,
        Dictionary<string, string> properties)
    {
        // Resolve paragraphs from parent path
        var paragraphs = ResolvePptParagraphsForFind(parentPath);
        if (paragraphs.Count == 0)
            throw new ArgumentException($"No paragraphs found at path: {parentPath}");

        // Support regex=true prop as alternative to r"..." prefix.
        // CONSISTENCY(find-regex): mirror of WordHandler.Set.cs:60-61. grep
        // "CONSISTENCY(find-regex)" for every project-wide call site.
        if (properties.TryGetValue("regex", out var regexFlag) && ParseHelpers.IsTruthySafe(regexFlag) && !findValue.StartsWith("r\"") && !findValue.StartsWith("r'"))
            findValue = $"r\"{findValue}\"";

        var (pattern, isRegex) = ParseFindPattern(findValue);

        // Find first match in any paragraph
        Drawing.Paragraph? targetPara = null;
        int splitPoint = -1;

        foreach (var para in paragraphs)
        {
            var runTexts = BuildPptRunTexts(para);
            if (runTexts.Count == 0) continue;
            var fullText = string.Concat(runTexts.Select(rt => rt.TextElement.Text));
            var matches = FindMatchRanges(fullText, pattern, isRegex);
            if (matches.Count > 0)
            {
                targetPara = para;
                var (matchStart, matchLen) = matches[0];
                splitPoint = isAfter ? matchStart + matchLen : matchStart;
                break;
            }
        }

        if (targetPara == null)
            throw new ArgumentException($"Text '{findValue}' not found in paragraphs at {parentPath}.");

        // Split run at the position
        var rts = BuildPptRunTexts(targetPara);
        Drawing.Run? insertAfterRun = null;

        foreach (var rt in rts)
        {
            if (splitPoint >= rt.Start && splitPoint <= rt.End)
            {
                if (splitPoint == rt.Start)
                    insertAfterRun = rt.Run.PreviousSibling<Drawing.Run>();
                else if (splitPoint == rt.End)
                    insertAfterRun = rt.Run;
                else
                {
                    SplitPptRunAtOffset(rt.Run, splitPoint - rt.Start);
                    insertAfterRun = rt.Run;
                }
                break;
            }
        }

        // Build and insert new run directly into targetPara (avoids path-based routing
        // that only supports /slide[N]/shape[M] paths, not table cell or other paths).
        var newRun = BuildPptRunFromProperties(properties);

        if (insertAfterRun != null)
            insertAfterRun.InsertAfterSelf(newRun);
        else
        {
            // Insert at beginning: before first run or end-paragraph props
            var firstChild = targetPara.FirstChild;
            if (firstChild != null)
                firstChild.InsertBeforeSelf(newRun);
            else
                targetPara.Append(newRun);
        }

        // Save all slides
        foreach (var slidePart in _doc.PresentationPart?.SlideParts ?? Enumerable.Empty<SlidePart>())
            slidePart.Slide?.Save();

        return parentPath;
    }

    /// <summary>
    /// Build a Drawing.Run from a properties dictionary (text, bold, italic, color, size, font, etc.)
    /// </summary>
    private static Drawing.Run BuildPptRunFromProperties(Dictionary<string, string> properties)
    {
        var newRun = new Drawing.Run();
        var rProps = new Drawing.RunProperties { Language = "en-US" };

        if (properties.TryGetValue("size", out var rSize))
            rProps.FontSize = (int)Math.Round(ParseFontSize(rSize) * 100);
        if (properties.TryGetValue("bold", out var rBold))
            rProps.Bold = IsTruthy(rBold);
        if (properties.TryGetValue("italic", out var rItalic))
            rProps.Italic = IsTruthy(rItalic);
        if (properties.TryGetValue("underline", out var rUnderline))
            rProps.Underline = rUnderline.ToLowerInvariant() switch
            {
                "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                "heavy" => Drawing.TextUnderlineValues.Heavy,
                "dotted" => Drawing.TextUnderlineValues.Dotted,
                "dash" => Drawing.TextUnderlineValues.Dash,
                "wavy" => Drawing.TextUnderlineValues.Wavy,
                "false" or "none" => Drawing.TextUnderlineValues.None,
                _ => throw new ArgumentException($"Invalid underline value: '{rUnderline}'.")
            };
        if (properties.TryGetValue("strikethrough", out var rStrike) || properties.TryGetValue("strike", out rStrike))
            rProps.Strike = rStrike.ToLowerInvariant() switch
            {
                "true" or "single" => Drawing.TextStrikeValues.SingleStrike,
                "double" => Drawing.TextStrikeValues.DoubleStrike,
                "false" or "none" => Drawing.TextStrikeValues.NoStrike,
                _ => throw new ArgumentException($"Invalid strikethrough value: '{rStrike}'.")
            };
        if (properties.TryGetValue("color", out var rColor))
            rProps.AppendChild(BuildSolidFill(rColor));
        if (properties.TryGetValue("font", out var rFont))
        {
            rProps.Append(new Drawing.LatinFont { Typeface = rFont });
            rProps.Append(new Drawing.EastAsianFont { Typeface = rFont });
        }
        if (properties.TryGetValue("spacing", out var rSpacing) || properties.TryGetValue("charspacing", out rSpacing))
            rProps.Spacing = (int)(ParseHelpers.SafeParseDouble(rSpacing, "charspacing") * 100);

        newRun.RunProperties = rProps;
        var runText = properties.GetValueOrDefault("text", "");
        newRun.Text = new Drawing.Text { Text = runText.Replace("\\n", "\n") };
        return newRun;
    }

    // ==================== Binary Extraction ====================
    //
    // Support for `officecli get --save <dest>`. The node's relId plus
    // the /slide[N]/ prefix in the path identifies the owning SlidePart;
    // the payload part is then looked up and its stream copied out.
    public bool TryExtractBinary(string path, string destPath, out string? contentType, out long byteCount)
    {
        contentType = null;
        byteCount = 0;
        var node = Get(path, 0);
        if (node == null) return false;
        if (!node.Format.TryGetValue("relId", out var relObj) || relObj is not string relId
            || string.IsNullOrEmpty(relId))
            return false;

        // Infer slide index from the path (/slide[N]/...).
        var m = System.Text.RegularExpressions.Regex.Match(path, @"^/slide\[(\d+)\]");
        if (!m.Success) return false;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return false;

        var slidePart = slideParts[slideIdx - 1];
        DocumentFormat.OpenXml.Packaging.OpenXmlPart? part = null;
        try { part = slidePart.GetPartById(relId); } catch { /* not on slide */ }
        if (part == null) return false;

        // BUG-R10-04: create the destination directory if missing so
        // `get --save ./outdir/file.bin` works when outdir doesn't exist.
        var destDir = Path.GetDirectoryName(destPath);
        if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            Directory.CreateDirectory(destDir);

        // CONSISTENCY(ole-cfb-wrap): unwrap CFB Ole10Native payload on read.
        byte[] rawBytes;
        using (var src = part.GetStream())
        using (var ms = new MemoryStream())
        {
            src.CopyTo(ms);
            rawBytes = ms.ToArray();
        }
        var payload = OfficeCli.Core.OleHelper.UnwrapOle10NativeIfCfb(rawBytes);
        File.WriteAllBytes(destPath, payload);
        byteCount = payload.Length;
        contentType = part.ContentType;
        return true;
    }

    // ==================== OLE Object Reading ====================
    //
    // Enumerate all OLE objects on a slide. PPTX wraps OLE in a
    // GraphicFrame whose GraphicData uri = "*/ole" contains a <p:oleObj>
    // element with progId + r:id. We walk descendants to catch both the
    // modern (p:oleObj as direct child) and alternate content fallback
    // forms. Orphan embedded parts (not referenced by any oleObj) are
    // surfaced the same way as the Excel reader, so nothing disappears.
    internal List<DocumentNode> CollectOleNodesForSlide(int slideNum, SlidePart slidePart)
    {
        var nodes = new List<DocumentNode>();
        var seenRelIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return nodes;

        // 1. Walk GraphicFrames hosting p:oleObj (strong-typed via SDK).
        int oleIdx = 0;
        foreach (var gf in shapeTree.Descendants<GraphicFrame>())
        {
            // A GraphicFrame may carry table/chart/ole — filter on the
            // presence of a strong-typed OleObject descendant.
            var oleObj = gf.Descendants<DocumentFormat.OpenXml.Presentation.OleObject>().FirstOrDefault();
            if (oleObj == null) continue;

            oleIdx++;
            var node = new DocumentNode
            {
                Path = $"/slide[{slideNum}]/ole[{oleIdx}]",
                Type = "ole",
                Text = oleObj.ProgId?.Value ?? "",
            };
            node.Format["objectType"] = "ole";
            if (oleObj.ProgId?.Value != null) node.Format["progId"] = oleObj.ProgId.Value;
            if (oleObj.Name?.Value != null) node.Format["name"] = oleObj.Name.Value;
            // CONSISTENCY(ole-display): always emit display key so callers can
            // rely on it being present; mirrors Word OLE DrawAspect normalization.
            node.Format["display"] = (oleObj.ShowAsIcon?.Value == true) ? "icon" : "content";
            // CONSISTENCY(ole-width-units): imgW/imgH (raw EMU) used to be
            // surfaced here but duplicated the unit-qualified width/height
            // emitted from the graphicFrame xfrm below. Kept internal only.

            // Extents from the frame's own xfrm.
            var xfrm = gf.Transform;
            if (xfrm?.Extents != null)
            {
                if (xfrm.Extents.Cx?.Value != null)
                    node.Format["width"] = OfficeCli.Core.EmuConverter.FormatEmu(xfrm.Extents.Cx.Value);
                if (xfrm.Extents.Cy?.Value != null)
                    node.Format["height"] = OfficeCli.Core.EmuConverter.FormatEmu(xfrm.Extents.Cy.Value);
            }

            var relId = oleObj.Id?.Value;
            if (!string.IsNullOrEmpty(relId))
            {
                node.Format["relId"] = relId;
                seenRelIds.Add(relId);
                try
                {
                    var part = slidePart.GetPartById(relId);
                    if (part != null)
                        OfficeCli.Core.OleHelper.PopulateFromPart(node, part, oleObj.ProgId?.Value);
                }
                catch
                {
                    // Ignore rel-join failures; keep whatever we got from XML.
                }
            }

            nodes.Add(node);
        }

        // CONSISTENCY(ole-orphan-indexing): orphan embedded parts are NOT
        // indexed under ole[N] to keep Get/Set/Remove in lockstep. Set/Remove
        // dispatch on schema-typed <p:oleObj> elements only; indexing orphans
        // here would produce Get-visible nodes that Set/Remove cannot
        // address. See ExcelHandler.Helpers.cs for the mirror comment.

        return nodes;
    }
}
