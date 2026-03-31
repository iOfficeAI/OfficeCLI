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
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        parentPath = NormalizeCellPath(parentPath);

        return type.ToLowerInvariant() switch
        {
            "slide" => AddSlide(parentPath, index, properties),
            "shape" or "textbox" when properties != null && properties.ContainsKey("formula") => AddEquation(parentPath, index, properties),
            "shape" or "textbox" => AddShape(parentPath, index, properties ?? new()),
            "picture" or "image" or "img" => AddPicture(parentPath, index, properties),
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
