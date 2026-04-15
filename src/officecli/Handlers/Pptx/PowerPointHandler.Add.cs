// Copyright 2025 OfficeCli (officecli.ai)
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
        properties = properties == null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : properties.Comparer == StringComparer.OrdinalIgnoreCase
                ? properties
                : new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);

        parentPath = NormalizeCellPath(parentPath);
        parentPath = ResolveIdPath(parentPath);

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
            "shape" or "textbox" => AddShape(parentPath, index, properties ?? new()),
            "picture" or "image" or "img" => AddPicture(parentPath, index, properties),
            "ole" or "oleobject" or "object" or "embed" => AddOle(parentPath, index, properties ?? new()),
            "chart" => AddChart(parentPath, index, properties),
            "table" => AddTable(parentPath, index, properties),
            "equation" or "formula" or "math" => AddEquation(parentPath, index, properties),
            "notes" or "note" => AddNotes(parentPath, index, properties),
            "video" or "audio" or "media" => AddMedia(parentPath, index, properties, type),
            "connector" or "connection" => AddConnector(parentPath, index, properties),
            "group" => AddGroup(parentPath, index, properties),
            "row" or "tr" => AddRow(parentPath, index, properties),
            "col" or "column" => AddColumn(parentPath, index, properties),
            "cell" or "tc" => AddCell(parentPath, index, properties),
            "animation" or "animate" => AddAnimation(parentPath, index, properties),
            "paragraph" or "para" => AddParagraph(parentPath, index, properties),
            "run" => AddRun(parentPath, index, properties),
            "zoom" or "slidezoom" or "slide-zoom" => AddZoom(parentPath, index, properties),
            "3dmodel" or "model3d" or "model" or "glb" => AddModel3D(parentPath, index, properties),
            _ => AddDefault(parentPath, index, properties, type)
        };
    }

}
