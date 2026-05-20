// Copyright 2025 OfficeCLI (officecli.ai)
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
            "table" => AddTable(parentPath, index, properties),
            "equation" or "formula" or "math" => AddEquation(parentPath, index, properties),
            "notes" or "note" => AddNotes(parentPath, index, properties),
            "video" or "audio" or "media" => AddMedia(parentPath, index, properties, type),
            "connector" or "connection" => AddConnector(parentPath, index, properties),
            "group" => AddGroup(parentPath, index, properties),
            "placeholder" or "ph" => AddPlaceholder(parentPath, index, properties),
            "row" or "tr" => AddRow(parentPath, index, properties),
            "col" or "column" => AddColumn(parentPath, index, properties),
            "cell" or "tc" => AddCell(parentPath, index, properties),
            "animation" or "animate" => AddAnimation(parentPath, index, properties),
            "paragraph" or "para" => AddParagraph(parentPath, index, properties),
            "run" => AddRun(parentPath, index, properties),
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
            _ => AddDefault(parentPath, index, properties, type)
        };
    }

}
