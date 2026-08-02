// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
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

    /// <summary>
    /// Scan all slides to initialize the global shape ID counter.
    /// Called once on document open (editable mode).
    /// </summary>
    private void InitShapeIdCounter() => InitShapeIdCounter(fromRawStreams: false);

    /// <param name="fromRawStreams">
    /// When true, scan each slide's cNvPr ids by reading its raw part stream
    /// instead of materializing the strongly-typed DOM. Touching
    /// <c>slidePart.Slide</c> (via <see cref="GetSlide"/>) loads the SDK root,
    /// and a later package Save then re-serializes that part — normalizing its
    /// lexical formatting and dropping any XML the SDK does not model, even on
    /// slides the command never targeted (#267). The open path passes true so a
    /// targeted edit leaves untouched slides byte-for-byte identical. The
    /// post-raw-set rebuild passes false: that path has just mutated a slide's
    /// live DOM whose new ids are not yet flushed to the stream, so it must read
    /// the DOM.
    /// </param>
    private void InitShapeIdCounter(bool fromRawStreams)
    {
        // CONSISTENCY(shape-id-high-range): auto-assigned ids start at 100000+
        // so they cannot collide with PowerPoint-authored ids (which sit in
        // the 1..99 range for placeholders and the 1000..99999 range for
        // regular shapes). This lets dump→replay preserve the original cNvPr
        // id verbatim for every shape (placeholder + regular) without risking
        // collision when a later mutation auto-assigns a fresh id.
        const uint minStartId = 100000;
        _usedShapeIds = new HashSet<uint>();
        uint maxId = minStartId - 1;

        foreach (var slidePart in GetSlideParts())
        {
            var ids = fromRawStreams
                ? ScanShapeIdsFromStream(slidePart)
                : ScanShapeIdsFromDom(slidePart);
            foreach (var id in ids)
            {
                _usedShapeIds.Add(id);
                if (id > maxId)
                    maxId = id;
            }
        }

        _nextShapeId = maxId + 1;
        if (_nextShapeId < maxId) // uint overflow
            _nextShapeId = minStartId;
    }

    private static IEnumerable<uint> ScanShapeIdsFromDom(SlidePart slidePart)
    {
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) yield break;
        foreach (var nvPr in shapeTree.Descendants<NonVisualDrawingProperties>())
            if (nvPr.Id?.HasValue == true)
                yield return nvPr.Id.Value;
    }

    // Scan cNvPr ids from the raw part XML without loading the strongly-typed
    // DOM, so the part stays clean and Save leaves it byte-for-byte intact
    // (#267). Matches any element with local name "cNvPr" regardless of
    // namespace — mirrors Descendants<NonVisualDrawingProperties>().
    private static IEnumerable<uint> ScanShapeIdsFromStream(OpenXmlPart part)
    {
        using var stream = part.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read);
        using var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings
        {
            DtdProcessing = System.Xml.DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            CloseInput = true,
        });
        while (reader.Read())
        {
            if (reader.NodeType == System.Xml.XmlNodeType.Element && reader.LocalName == "cNvPr")
            {
                var idStr = reader.GetAttribute("id");
                if (idStr != null && uint.TryParse(idStr, out var id))
                    yield return id;
            }
        }
    }

    /// <summary>
    /// Return true if <paramref name="id"/> is already claimed by any cNvPr in
    /// the given shapeTree, or globally in <see cref="_usedShapeIds"/>.
    /// </summary>
    private bool ShapeIdInUse(ShapeTree shapeTree, uint id)
    {
        if (_usedShapeIds != null && _usedShapeIds.Contains(id))
            return true;
        if (shapeTree != null)
        {
            foreach (var nvPr in shapeTree.Descendants<NonVisualDrawingProperties>())
            {
                if (nvPr.Id?.HasValue == true && nvPr.Id.Value == id)
                    return true;
            }
        }
        return false;
    }

    /// <summary>
    /// CONSISTENCY(dump-replay-id): honor a caller-supplied "id" property so
    /// that dump→batch round-trip preserves @id=N references; mirrors docx
    /// Add.Structure.cs:1118 for numbering ids. id=0 / non-numeric / missing
    /// → auto-assign via <see cref="GenerateUniqueShapeId"/>. Collisions with
    /// an in-use id throw rather than silently renumber.
    /// </summary>
    private uint AcquireShapeId(ShapeTree shapeTree, Dictionary<string, string> properties)
    {
        if (properties != null
            && properties.TryGetValue("id", out var idStr)
            && uint.TryParse(idStr, out var requestedId)
            && requestedId > 0)
        {
            // CONSISTENCY(per-slide-id-scope): OOXML cNvPr ids only need to
            // be unique within a slide — PowerPoint authors ids 1/2/3/...
            // starting fresh on every slide. The global _usedShapeIds set
            // tracks ids across slides for the auto-assign path (to keep
            // animation spid references stable across mutations), but for
            // a caller-supplied id (dump→replay round-trip) the relevant
            // scope is the parent shapeTree. Without this, dump emitted
            // id=2 on every slide would error from slide 2 onward.
            //
            // CONSISTENCY(sptree-root-id-not-a-sibling): the shapeTree's own
            // <p:nvGrpSpPr><p:cNvPr id="1"/></p:nvGrpSpPr> is the wrapper, not
            // a child shape. PowerPoint routinely authors a title placeholder
            // with id=1 alongside the spTree root id=1 (see e.g. video.pptx
            // slide2 — `<p:cNvPr id="1" name="">` for the group AND
            // `<p:cNvPr id="1" name="Title">` for the title sp). Treating the
            // root id as a collision blocks legitimate dump→replay of any
            // such PowerPoint-authored slide.
            if (shapeTree != null)
            {
                var rootNvPr = shapeTree.GetFirstChild<NonVisualGroupShapeProperties>()
                    ?.GetFirstChild<NonVisualDrawingProperties>();
                foreach (var nvPr in shapeTree.Descendants<NonVisualDrawingProperties>())
                {
                    if (ReferenceEquals(nvPr, rootNvPr)) continue;
                    if (nvPr.Id?.HasValue == true && nvPr.Id.Value == requestedId)
                        throw new ArgumentException(
                            $"id {requestedId} already in use in this shapeTree. " +
                            "Use a different id or omit to auto-assign.");
                }
            }
            _usedShapeIds?.Add(requestedId);
            if (requestedId >= _nextShapeId)
                _nextShapeId = requestedId + 1;
            return requestedId;
        }
        return GenerateUniqueShapeId(shapeTree);
    }

    /// <summary>
    /// Generate a unique deterministic cNvPr.Id across all slides.
    /// Uses global instance counter for reproducible, non-repeating IDs.
    /// </summary>
    private uint GenerateUniqueShapeId(ShapeTree shapeTree)
    {
        // See CONSISTENCY(shape-id-high-range) in InitShapeIdCounter.
        const uint minStartId = 100000;
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
    /// Get the cNvPr container (NonVisualDrawingProperties) for an element,
    /// erased to OpenXmlElement so callers can read extLst across both
    /// PresentationML (sp/pic/cxnSp) and DrawingML (graphicFrame) variants.
    /// Mirrors GetCNvPrId across all element types so extLst content
    /// (creationId, ...) can be read from a single accessor.
    /// </summary>
    internal static OpenXmlElement? GetCNvPr(OpenXmlElement element)
    {
        return element switch
        {
            Shape s => (OpenXmlElement?)s.NonVisualShapeProperties?.NonVisualDrawingProperties,
            Picture p => (OpenXmlElement?)p.NonVisualPictureProperties?.NonVisualDrawingProperties,
            GraphicFrame gf => (OpenXmlElement?)gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties,
            ConnectionShape c => (OpenXmlElement?)c.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties,
            GroupShape g => (OpenXmlElement?)g.NonVisualGroupShapeProperties?.NonVisualDrawingProperties,
            _ => null
        };
    }

    // bt-3: PowerPoint 2015+ stamps every shape with a stable creationId GUID
    // stored inside cNvPr's extLst (ext uri "{FF2B5EF4-...}" → p15:creationId).
    // Without surfacing this, dump→replay loses modern shape identity used by
    // change tracking, comments, and timeline anchors.
    private const string CreationIdExtUri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}";

    internal static string? ReadCNvPrCreationId(OpenXmlElement? element)
    {
        if (element == null) return null;
        var cNvPr = GetCNvPr(element);
        if (cNvPr == null) return null;
        // Walk children by local name — the extLst element type differs
        // between PresentationML and DrawingML cNvPr variants, and the
        // creationId child sits in the p15 namespace. Local-name walk
        // sidesteps the namespace zoo.
        foreach (var child in cNvPr.ChildElements)
        {
            if (child.LocalName != "extLst") continue;
            foreach (var ext in child.ChildElements)
            {
                if (ext.LocalName != "ext") continue;
                string? uri = null;
                foreach (var a in ext.GetAttributes())
                {
                    if (a.LocalName == "uri") { uri = a.Value; break; }
                }
                if (uri == null || !uri.Equals(CreationIdExtUri, StringComparison.OrdinalIgnoreCase))
                    continue;
                foreach (var inner in ext.ChildElements)
                {
                    if (inner.LocalName != "creationId") continue;
                    foreach (var a in inner.GetAttributes())
                    {
                        if (a.LocalName == "val") return a.Value;
                    }
                }
            }
        }
        return null;
    }
}
