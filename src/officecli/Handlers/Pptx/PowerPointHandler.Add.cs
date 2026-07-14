// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    public string Add(string parentPath, string type, InsertPosition? position, Dictionary<string, string> properties)
    {
        Modified = true;
        LastUnrecognizedLatex = new List<string>();
        // CONSISTENCY(prop-key-case): property keys are case-insensitive
        // ("SRC"/"src"/"Src" all resolve the same). Normalize once at the
        // dispatch entry so every AddXxx helper can rely on TryGetValue("src").
        properties = properties switch
        {
            null => new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
            // Preserve TrackingPropertyDictionary so handler-as-truth read
            // tracking survives the entry normalization. The tracking
            // comparer wraps OrdinalIgnoreCase so case-insensitive lookup
            // works as intended.
            OfficeCli.Core.TrackingPropertyDictionary => properties,
            var p when p.Comparer == StringComparer.OrdinalIgnoreCase => p,
            _ => new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase),
        };

        parentPath = NormalizePptxPathSegmentCasing(parentPath);
        parentPath = NormalizeCellPath(parentPath);
        parentPath = ResolveIdPath(parentPath);
        parentPath = ResolveLastPredicates(parentPath);

        // Reject negative --index up front with a clean message instead of
        // letting it fall through and surface as a raw .NET
        // ArgumentOutOfRangeException from collection indexing. Matches
        // the guard in WordHandler.Add.
        if (position?.Index.HasValue == true && position.Index.Value < 0)
            throw new ArgumentException("--index must be non-negative.");

        // Resolve --after/--before to index (handles find: prefix)
        var index = ResolveAnchorPosition(parentPath, position);

        // Handle find: prefix — text-based anchoring in PPT paragraphs
        if (index == FindAnchorIndex && position != null)
        {
            var anchorValue = (position.After ?? position.Before)!;
            var findValue = anchorValue["find:".Length..];
            var isAfter = position.After != null;
            return AddPptAtFindPosition(parentPath, type, findValue, isAfter, properties);
        }

        return type.ToLowerInvariant() switch
        {
            "slide" => AddSlide(parentPath, index, properties),
            "shape" or "textbox" when properties != null && properties.ContainsKey("formula") => AddEquation(parentPath, index, properties),
            // Forward the requested element type so AddShape can distinguish
            // `--type shape` (geometry shape) from `--type textbox` (writes
            // <p:cNvSpPr txBox="1"/>) even when neither geometry nor text props
            // are supplied. The dump emitter splits text into separate
            // paragraph/run adds, so the AddShape call carries no `text=` —
            // without this hint, replay can't tell the two flavors apart.
            "shape" or "textbox" => AddShape(parentPath, index, properties ?? new(), type.ToLowerInvariant()),
            "picture" or "image" or "img" => AddPicture(parentPath, index, properties),
            "ole" or "oleobject" or "object" or "embed" => AddOle(parentPath, index, properties ?? new()),
            "chart" => AddChart(parentPath, index, properties),
            // CONSISTENCY(chart-series-alias): schema element name is
            // "chart-series" (elementAliases: ["series"]) — accept both.
            "series" or "chart-series" or "chartseries" => AddChartSeries(parentPath, properties ?? new()),
            "table" => AddTable(parentPath, index, properties),
            "equation" or "formula" or "math" => AddEquation(parentPath, index, properties),
            "diagram" or "flowchart" => AddDiagram(parentPath, index, properties ?? new()),
            "notes" or "note" => AddNotes(parentPath, index, properties),
            "video" or "audio" or "media" => AddMedia(parentPath, index, properties, type),
            "connector" or "connection" => AddConnector(parentPath, index, properties),
            "group" => AddGroup(parentPath, index, properties),
            "placeholder" or "ph" => AddPlaceholder(parentPath, index, properties),
            "row" or "tr" => AddRow(parentPath, index, properties),
            "col" or "column" => AddColumn(parentPath, index, properties),
            "cell" or "tc" => AddCell(parentPath, index, properties),
            "animation" or "animate" => AddAnimation(parentPath, index, properties),
            // CONSISTENCY(hyperlink-shape-parent): `add --type hyperlink /slide[N]/shape[M]`
            // attaches an action hyperlink to an existing shape. ResolveLogicalPath only
            // covers /slide[N]/{table,placeholder}[X]; shape parents fall to a generic
            // XML-localName navigator that doesn't know <p:sp>, so the dispatch needs
            // its own entry. Mirrors AddShape's `link=` branch.
            "hyperlink" or "hlink" => AddHyperlinkOnShape(parentPath, properties),
            "paragraph" or "para" => AddParagraph(parentPath, index, properties),
            "run" => AddRun(parentPath, index, properties),
            // CONSISTENCY(linebreak-shape-path): `add --type linebreak /slide[N]/shape[M]/paragraph[K]`
            // inserts an <a:br/> into the paragraph. Same dispatch-gap class as AddRun/hyperlink —
            // without an explicit case, the type falls to AddDefault → ResolveLogicalPath (no shape
            // support) → generic XML fallback by localName "shape" misses <p:sp> → "parent not found".
            "linebreak" or "br" or "line-break" => AddLineBreak(parentPath, index, properties),
            "zoom" or "slidezoom" or "slide-zoom" => AddZoom(parentPath, index, properties),
            "3dmodel" or "model3d" or "model" or "glb" => AddModel3D(parentPath, index, properties),
            // BUG-R36-B11: legacy slide comments lifecycle.
            "comment" or "note-comment" => AddSlideComment(parentPath, index, properties),
            // Modern p188 (Office 2018/8) threaded comments — distinct OOXML
            // element living in PowerPointCommentPart (/ppt/comments/…). Top-
            // level threads and replies share the dispatch (parent= prop
            // discriminates).
            "moderncomment" or "modern-comment" or "thread" or "threadedcomment"
                => AddModernComment(parentPath, index, properties),
            // A slide transition is a per-slide property, not a child element. Without
            // an explicit case it would fall to AddDefault → GenericXmlQuery and write a
            // bare <p:transition/> that the transition Get path can't resolve (phantom
            // element, false success). Reject with a redirect to the Set API.
            "transition" => throw new ArgumentException(
                "Transition is a slide property, not an element type. " +
                "Use: Set /slide[N] --prop transition=<effect> (e.g. transition=fade) to set a slide transition."),
            _ => AddDefault(parentPath, index, properties, type)
        };
    }

}
