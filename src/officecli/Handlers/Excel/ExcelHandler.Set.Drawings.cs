// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeCli.Handlers;

// Per-element-type Set helpers for drawing/anchor paths (sparkline, ole,
// picture, shape, slicer). Mechanically extracted from the original
// god-method Set(); each helper owns one path-pattern's full handling.
public partial class ExcelHandler
{
    private List<string> SetSparklineByPath(Match m, Dictionary<string, string> properties)
    {
        var spkSheet = m.Groups[1].Value;
        var spkIdx = int.Parse(m.Groups[2].Value);
        var spkWorksheet = FindWorksheet(spkSheet) ?? throw SheetNotFoundException(spkSheet);
        var spkGroup = GetSparklineGroup(spkWorksheet, spkIdx)
            ?? throw new ArgumentException($"Sparkline[{spkIdx}] not found in sheet '{spkSheet}'");

        var unsup = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "type":
                    spkGroup.Type = value.ToLowerInvariant() switch
                    {
                        "column" => X14.SparklineTypeValues.Column,
                        "stacked" => X14.SparklineTypeValues.Stacked,
                        _ => null
                    };
                    break;
                case "color":
                    spkGroup.SeriesColor = new X14.SeriesColor { Rgb = ParseHelpers.NormalizeArgbColor(value) };
                    break;
                case "negativecolor":
                    spkGroup.NegativeColor = new X14.NegativeColor { Rgb = ParseHelpers.NormalizeArgbColor(value) };
                    break;
                case "markers":
                    spkGroup.Markers = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "highpoint":
                    spkGroup.High = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "lowpoint":
                    spkGroup.Low = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "firstpoint":
                    spkGroup.First = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "lastpoint":
                    spkGroup.Last = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "negative":
                    spkGroup.Negative = ParseHelpers.IsTruthy(value) ? (bool?)true : null;
                    break;
                case "lineweight":
                    if (double.TryParse(value, out var lw)) spkGroup.LineWeight = lw;
                    break;
                case "datarange" or "range":
                {
                    var newRangeRef = value.Contains('!') ? value : $"{spkSheet}!{value}";
                    foreach (var spk in spkGroup.Descendants<X14.Sparkline>())
                    {
                        var f = spk.GetFirstChild<DocumentFormat.OpenXml.Office.Excel.Formula>();
                        if (f != null) f.Text = newRangeRef;
                        else spk.InsertAt(new DocumentFormat.OpenXml.Office.Excel.Formula(newRangeRef), 0);
                    }
                    break;
                }
                case "location" or "cell":
                {
                    foreach (var spk in spkGroup.Descendants<X14.Sparkline>())
                    {
                        var r = spk.GetFirstChild<DocumentFormat.OpenXml.Office.Excel.ReferenceSequence>();
                        if (r != null) r.Text = value;
                        else spk.AppendChild(new DocumentFormat.OpenXml.Office.Excel.ReferenceSequence(value));
                    }
                    break;
                }
                default:
                    unsup.Add(key);
                    break;
            }
        }
        SaveWorksheet(spkWorksheet);
        return unsup;
    }

    private List<string> SetOleByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var oleIdxSet = int.Parse(m.Groups[1].Value);
        var oleWs = GetSheet(worksheet);
        var oleElements = oleWs.Descendants<OleObject>().ToList();
        if (oleIdxSet < 1 || oleIdxSet > oleElements.Count)
            throw new ArgumentException($"OLE object index {oleIdxSet} out of range (1..{oleElements.Count})");
        var oleObjSet = oleElements[oleIdxSet - 1];
        var oleUnsupportedSet = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "path" or "src":
                {
                    if (oleObjSet.Id?.Value is string oldRel && !string.IsNullOrEmpty(oldRel))
                    {
                        try { worksheet.DeletePart(oldRel); } catch { }
                    }
                    var (newRel, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(worksheet, value, _filePath);
                    oleObjSet.Id = newRel;
                    if (!properties.ContainsKey("progId") && !properties.ContainsKey("progid"))
                    {
                        var autoProgId = OfficeCli.Core.OleHelper.DetectProgId(value);
                        OfficeCli.Core.OleHelper.ValidateProgId(autoProgId);
                        oleObjSet.ProgId = autoProgId;
                    }
                    break;
                }
                case "progid":
                    OfficeCli.Core.OleHelper.ValidateProgId(value);
                    oleObjSet.ProgId = value;
                    break;
                case "display":
                    // CONSISTENCY(excel-ole-display): Excel Add rejects 'display'
                    // with ArgumentException; Set must do the same instead of
                    // falling into the default unsupported branch.
                    throw new ArgumentException(
                        "'display' property is not supported for Excel OLE "
                        + "(Excel always shows objects as icon). Remove --prop display.");
                case "width":
                case "height":
                {
                    // CONSISTENCY(ole-width-units): accept either bare integer cell-span or unit-qualified size.
                    long emuTotal;
                    try { emuTotal = ParseAnchorDimensionEmu(value, key.ToLowerInvariant()); }
                    catch { oleUnsupportedSet.Add(key); break; }
                    if (emuTotal < 0) { oleUnsupportedSet.Add(key); break; }
                    var objectPrSet = oleObjSet.GetFirstChild<EmbeddedObjectProperties>();
                    var objAnchorSet = objectPrSet?.GetFirstChild<ObjectAnchor>();
                    var fromMSet = objAnchorSet?.GetFirstChild<FromMarker>();
                    var toMSet = objAnchorSet?.GetFirstChild<ToMarker>();
                    if (fromMSet == null || toMSet == null) { oleUnsupportedSet.Add(key); break; }
                    if (key.Equals("width", StringComparison.OrdinalIgnoreCase))
                    {
                        int.TryParse(fromMSet.GetFirstChild<XDR.ColumnId>()?.Text ?? "0", out var fromCol);
                        long.TryParse(fromMSet.GetFirstChild<XDR.ColumnOffset>()?.Text ?? "0", out var fromColOff);
                        long wholeCols = emuTotal / EmuPerColApprox;
                        long remCols = emuTotal % EmuPerColApprox;
                        var toColChild = toMSet.GetFirstChild<XDR.ColumnId>();
                        if (toColChild != null) toColChild.Text = (fromCol + (int)wholeCols).ToString();
                        var toColOffChild = toMSet.GetFirstChild<XDR.ColumnOffset>();
                        if (toColOffChild != null) toColOffChild.Text = (fromColOff + remCols).ToString();
                        else toMSet.InsertAfter(new XDR.ColumnOffset((fromColOff + remCols).ToString()), toColChild);
                    }
                    else
                    {
                        int.TryParse(fromMSet.GetFirstChild<XDR.RowId>()?.Text ?? "0", out var fromRow);
                        long.TryParse(fromMSet.GetFirstChild<XDR.RowOffset>()?.Text ?? "0", out var fromRowOff);
                        long wholeRows = emuTotal / EmuPerRowApprox;
                        long remRows = emuTotal % EmuPerRowApprox;
                        var toRowChild = toMSet.GetFirstChild<XDR.RowId>();
                        if (toRowChild != null) toRowChild.Text = (fromRow + (int)wholeRows).ToString();
                        var toRowOffChild = toMSet.GetFirstChild<XDR.RowOffset>();
                        if (toRowOffChild != null) toRowOffChild.Text = (fromRowOff + remRows).ToString();
                        else toMSet.InsertAfter(new XDR.RowOffset((fromRowOff + remRows).ToString()), toRowChild);
                    }
                    break;
                }
                case "anchor":
                {
                    // CONSISTENCY(ole-width-units): mirror Add-side warn — width/height
                    // dropped silently when anchor= present.
                    if (properties.ContainsKey("width") || properties.ContainsKey("height"))
                        Console.Error.WriteLine(
                            "Warning: 'width'/'height' are ignored when 'anchor' is provided (anchor defines the full rectangle).");
                    var anchorM = Regex.Match(value ?? "", @"^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$", RegexOptions.IgnoreCase);
                    if (!anchorM.Success) { oleUnsupportedSet.Add(key); break; }
                    var objectPrAnc = oleObjSet.GetFirstChild<EmbeddedObjectProperties>();
                    var objAnchorAnc = objectPrAnc?.GetFirstChild<ObjectAnchor>();
                    var fromMAnc = objAnchorAnc?.GetFirstChild<FromMarker>();
                    var toMAnc = objAnchorAnc?.GetFirstChild<ToMarker>();
                    if (fromMAnc == null || toMAnc == null) { oleUnsupportedSet.Add(key); break; }
                    int newFromCol = ColumnNameToIndex(anchorM.Groups[1].Value) - 1;
                    int newFromRow = int.Parse(anchorM.Groups[2].Value) - 1;
                    int newToCol, newToRow;
                    if (anchorM.Groups[3].Success)
                    {
                        newToCol = ColumnNameToIndex(anchorM.Groups[3].Value) - 1;
                        newToRow = int.Parse(anchorM.Groups[4].Value) - 1;
                    }
                    else
                    {
                        newToCol = newFromCol + 2;
                        newToRow = newFromRow + 3;
                    }
                    var fromColChild = fromMAnc.GetFirstChild<XDR.ColumnId>();
                    if (fromColChild != null) fromColChild.Text = newFromCol.ToString();
                    var fromRowChild = fromMAnc.GetFirstChild<XDR.RowId>();
                    if (fromRowChild != null) fromRowChild.Text = newFromRow.ToString();
                    var fromColOffChild = fromMAnc.GetFirstChild<XDR.ColumnOffset>();
                    if (fromColOffChild != null) fromColOffChild.Text = "0";
                    var fromRowOffChild = fromMAnc.GetFirstChild<XDR.RowOffset>();
                    if (fromRowOffChild != null) fromRowOffChild.Text = "0";
                    var toColChildAnc = toMAnc.GetFirstChild<XDR.ColumnId>();
                    if (toColChildAnc != null) toColChildAnc.Text = newToCol.ToString();
                    var toRowChildAnc = toMAnc.GetFirstChild<XDR.RowId>();
                    if (toRowChildAnc != null) toRowChildAnc.Text = newToRow.ToString();
                    var toColOffChildAnc = toMAnc.GetFirstChild<XDR.ColumnOffset>();
                    if (toColOffChildAnc != null) toColOffChildAnc.Text = "0";
                    var toRowOffChildAnc = toMAnc.GetFirstChild<XDR.RowOffset>();
                    if (toRowOffChildAnc != null) toRowOffChildAnc.Text = "0";
                    break;
                }
                default:
                    oleUnsupportedSet.Add(key);
                    break;
            }
        }
        SaveWorksheet(worksheet);
        return oleUnsupportedSet;
    }

    private List<string> SetPictureByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var picIdx = int.Parse(m.Groups[1].Value);
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/pictures");
        var wsDrawing = drawingsPart.WorksheetDrawing
            ?? throw new ArgumentException("Sheet has no drawings/pictures");

        var picAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Picture>().Any()).ToList();
        if (picIdx < 1 || picIdx > picAnchors.Count)
            throw new ArgumentException($"Picture index {picIdx} out of range (1..{picAnchors.Count})");

        var anchor = picAnchors[picIdx - 1];
        var picUnsupported = new List<string>();

        // CONSISTENCY(picture-crop): mirror Add — accept crop.l/r/t/b,
        // srcRect=l=..,r=..,t=..,b=.., and cropLeft/Right/Top/Bottom keys.
        // ParseSrcRect builds a Drawing.SourceRectangle from any subset.
        // We collect crop keys here and apply once after the property loop
        // so multiple crop keys in one Set call merge instead of clobber.
        var cropProps = new Dictionary<string, string>();
        var cropKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "crop.l", "crop.r", "crop.t", "crop.b",
            "srcRect", "cropLeft", "cropRight", "cropTop", "cropBottom"
        };

        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (cropKeys.Contains(key)) { cropProps[key] = value; continue; }
            if (TrySetAnchorPosition(anchor, lk, value)) continue;

            var spPr = anchor.Descendants<XDR.ShapeProperties>().FirstOrDefault();
            if (TrySetRotation(spPr, lk, value)) continue;
            if (TrySetShapeFlip(spPr, lk, value)) continue;
            if (TrySetShapeEffect(spPr, lk, value)) continue;

            switch (lk)
            {
                case "alt":
                    var nvProps = anchor.Descendants<XDR.NonVisualDrawingProperties>().FirstOrDefault();
                    if (nvProps != null) nvProps.Description = value;
                    break;
                default:
                    picUnsupported.Add(key);
                    break;
            }
        }

        if (cropProps.Count > 0)
        {
            var picture = anchor.Descendants<XDR.Picture>().FirstOrDefault();
            var blipFill = picture?.BlipFill;
            if (blipFill != null)
            {
                var newSrcRect = ParseSrcRect(cropProps);
                // Replace any existing <a:srcRect> with the new one. If
                // ParseSrcRect returns null (no valid crop values), drop the
                // existing srcRect entirely so the XML stays clean.
                foreach (var existing in blipFill.Elements<Drawing.SourceRectangle>().ToList())
                    existing.Remove();
                if (newSrcRect != null)
                {
                    // CONSISTENCY(ooxml-element-order): srcRect must precede
                    // the fill-mode element (stretch/tile) inside blipFill.
                    var fillMode = (OpenXmlElement?)blipFill.GetFirstChild<Drawing.Stretch>()
                        ?? blipFill.GetFirstChild<Drawing.Tile>();
                    if (fillMode != null)
                        blipFill.InsertBefore(newSrcRect, fillMode);
                    else
                        blipFill.AppendChild(newSrcRect);
                }
            }
            else
            {
                foreach (var k in cropProps.Keys) picUnsupported.Add(k);
            }
        }

        drawingsPart.WorksheetDrawing.Save();
        return picUnsupported;
    }

    private List<string> SetShapeByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var shpIdx = int.Parse(m.Groups[1].Value);
        var drawingsPart = worksheet.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/shapes");
        var wsDrawing = drawingsPart.WorksheetDrawing
            ?? throw new ArgumentException("Sheet has no drawings/shapes");

        var shpAnchors = wsDrawing.Elements<XDR.TwoCellAnchor>()
            .Where(a => a.Descendants<XDR.Shape>().Any()).ToList();
        if (shpIdx < 1 || shpIdx > shpAnchors.Count)
            throw new ArgumentException($"Shape index {shpIdx} out of range (1..{shpAnchors.Count})");

        var anchor = shpAnchors[shpIdx - 1];
        var shape = anchor.Descendants<XDR.Shape>().First();
        var shpUnsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (TrySetAnchorPosition(anchor, lk, value)) continue;
            if (TrySetRotation(shape.ShapeProperties, lk, value)) continue;
            if (TrySetShapeFlip(shape.ShapeProperties, lk, value)) continue;
            if (TrySetShapeFontProp(shape, lk, value)) continue;

            // For effects on shapes: check if fill=none → text-level, otherwise shape-level
            if (lk is "shadow" or "glow" or "reflection" or "softedge")
            {
                var spPr = shape.ShapeProperties;
                if (spPr == null) continue;
                var isNoFill = spPr.GetFirstChild<Drawing.NoFill>() != null;
                var normalizedVal = value.Replace(':', '-');

                if (isNoFill && lk is "shadow" or "glow")
                {
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        if (lk == "shadow")
                            OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.OuterShadow>(run, normalizedVal, () =>
                                OfficeCli.Core.DrawingEffectsHelper.BuildOuterShadow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                        else
                            OfficeCli.Core.DrawingEffectsHelper.ApplyTextEffect<Drawing.Glow>(run, normalizedVal, () =>
                                OfficeCli.Core.DrawingEffectsHelper.BuildGlow(normalizedVal, OfficeCli.Core.DrawingEffectsHelper.BuildRgbColor));
                    }
                }
                else
                {
                    TrySetShapeEffect(spPr, lk, value);
                }
                continue;
            }

            switch (lk)
            {
                case "name":
                {
                    var nvProps = shape.NonVisualShapeProperties?.GetFirstChild<XDR.NonVisualDrawingProperties>();
                    if (nvProps != null) nvProps.Name = value;
                    break;
                }
                case "text":
                {
                    var txBody = shape.TextBody;
                    if (txBody != null)
                    {
                        var firstPara = txBody.Elements<Drawing.Paragraph>().FirstOrDefault();
                        var pProps = firstPara?.ParagraphProperties?.CloneNode(true);
                        var rProps = firstPara?.Elements<Drawing.Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                        txBody.RemoveAllChildren<Drawing.Paragraph>();
                        var lines = value.Replace("\\n", "\n").Split('\n');
                        foreach (var line in lines)
                        {
                            var para = new Drawing.Paragraph();
                            if (pProps != null) para.AppendChild(pProps.CloneNode(true));
                            var run = new Drawing.Run(new Drawing.Text(line));
                            if (rProps != null) run.RunProperties = (Drawing.RunProperties)rProps.CloneNode(true);
                            para.AppendChild(run);
                            txBody.AppendChild(para);
                        }
                    }
                    break;
                }
                case "font":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.RemoveAllChildren<Drawing.LatinFont>();
                        rPr.RemoveAllChildren<Drawing.EastAsianFont>();
                        rPr.AppendChild(new Drawing.LatinFont { Typeface = value });
                        rPr.AppendChild(new Drawing.EastAsianFont { Typeface = value });
                    }
                    break;
                case "size":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.FontSize = (int)Math.Round(ParseHelpers.ParseFontSize(value) * 100);
                    }
                    break;
                case "bold":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.Bold = IsTruthy(value);
                    }
                    break;
                case "italic":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.Italic = IsTruthy(value);
                    }
                    break;
                case "color":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.RemoveAllChildren<Drawing.SolidFill>();
                        var (cRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                        OfficeCli.Core.DrawingEffectsHelper.InsertFillInRunProperties(rPr,
                            new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = cRgb }));
                    }
                    break;
                case "underline":
                    foreach (var run in shape.Descendants<Drawing.Run>())
                    {
                        var rPr = run.RunProperties ?? (run.RunProperties = new Drawing.RunProperties());
                        rPr.Underline = value.ToLowerInvariant() switch
                        {
                            "true" or "single" or "sng" => Drawing.TextUnderlineValues.Single,
                            "double" or "dbl" => Drawing.TextUnderlineValues.Double,
                            "heavy" => Drawing.TextUnderlineValues.Heavy,
                            "dotted" => Drawing.TextUnderlineValues.Dotted,
                            "dash" => Drawing.TextUnderlineValues.Dash,
                            "wavy" => Drawing.TextUnderlineValues.Wavy,
                            "false" or "none" => Drawing.TextUnderlineValues.None,
                            _ => throw new ArgumentException($"Invalid underline value: '{value}'. Valid values: single, double, heavy, dotted, dash, wavy, none.")
                        };
                    }
                    break;
                case "fill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr != null)
                    {
                        spPr.RemoveAllChildren<Drawing.SolidFill>();
                        spPr.RemoveAllChildren<Drawing.NoFill>();
                        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                            spPr.AppendChild(new Drawing.NoFill());
                        else
                        {
                            var (fRgb, _) = ParseHelpers.SanitizeColorForOoxml(value);
                            spPr.AppendChild(new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = fRgb }));
                        }
                    }
                    break;
                }
                case "align":
                    foreach (var para in shape.Descendants<Drawing.Paragraph>())
                    {
                        var pPr = para.ParagraphProperties ?? (para.ParagraphProperties = new Drawing.ParagraphProperties());
                        pPr.Alignment = value.ToLowerInvariant() switch
                        {
                            "center" or "c" or "ctr" => Drawing.TextAlignmentTypeValues.Center,
                            "right" or "r" => Drawing.TextAlignmentTypeValues.Right,
                            "justify" or "justified" or "j" => Drawing.TextAlignmentTypeValues.Justified,
                            "left" or "l" => Drawing.TextAlignmentTypeValues.Left,
                            _ => throw new ArgumentException($"Invalid align value: '{value}'. Valid values: left, center, right, justify.")
                        };
                    }
                    break;
                case "valign":
                {
                    var txBody = shape.TextBody;
                    var bodyPr = txBody?.GetFirstChild<Drawing.BodyProperties>();
                    if (bodyPr != null)
                    {
                        bodyPr.Anchor = value.ToLowerInvariant() switch
                        {
                            "top" or "t" => Drawing.TextAnchoringTypeValues.Top,
                            "center" or "ctr" or "middle" or "m" or "c" => Drawing.TextAnchoringTypeValues.Center,
                            "bottom" or "b" => Drawing.TextAnchoringTypeValues.Bottom,
                            _ => throw new ArgumentException($"Invalid valign value: '{value}'. Valid values: top, center, bottom.")
                        };
                    }
                    break;
                }
                case "gradientfill":
                {
                    var spPr = shape.ShapeProperties;
                    if (spPr != null)
                    {
                        spPr.RemoveAllChildren<Drawing.SolidFill>();
                        spPr.RemoveAllChildren<Drawing.NoFill>();
                        spPr.RemoveAllChildren<Drawing.GradientFill>();
                        // CONSISTENCY(shape-gradient-fill): reuse Add-branch parser.
                        spPr.AppendChild(BuildShapeGradientFill(value));
                    }
                    break;
                }
                case "line" or "border":
                {
                    // CONSISTENCY(shape-line): mirror Add — accept "none" or "color[:width[:style]]".
                    var spPr = shape.ShapeProperties;
                    if (spPr == null) break;
                    spPr.RemoveAllChildren<Drawing.Outline>();
                    if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        spPr.AppendChild(new Drawing.Outline(new Drawing.NoFill()));
                        break;
                    }
                    var parts = value.Split(':');
                    var (lRgb, _) = ParseHelpers.SanitizeColorForOoxml(parts[0]);
                    var outline = new Drawing.Outline(
                        new Drawing.SolidFill(new Drawing.RgbColorModelHex { Val = lRgb }));
                    if (parts.Length > 1
                        && double.TryParse(parts[1], System.Globalization.NumberStyles.Float,
                            System.Globalization.CultureInfo.InvariantCulture, out var wpt))
                    {
                        outline.Width = (int)Math.Round(wpt * 12700);
                    }
                    if (parts.Length > 2)
                    {
                        var dash = parts[2].ToLowerInvariant() switch
                        {
                            "dash" => Drawing.PresetLineDashValues.Dash,
                            "dot" => Drawing.PresetLineDashValues.Dot,
                            "dashdot" => Drawing.PresetLineDashValues.DashDot,
                            "longdash" => Drawing.PresetLineDashValues.LargeDash,
                            "solid" => Drawing.PresetLineDashValues.Solid,
                            _ => (Drawing.PresetLineDashValues?)null
                        };
                        if (dash != null)
                            outline.AppendChild(new Drawing.PresetDash { Val = dash });
                    }
                    spPr.AppendChild(outline);
                    break;
                }
                case "alt" or "alttext" or "descr" or "description":
                {
                    var altNv = shape.NonVisualShapeProperties?
                        .GetFirstChild<XDR.NonVisualDrawingProperties>();
                    if (altNv != null) altNv.Description = value;
                    break;
                }
                case "margin":
                {
                    // CONSISTENCY(shape-margin): mirror Add — margin is text-body
                    // inset in points, applied to all four sides equally.
                    var bodyPr = shape.TextBody?.GetFirstChild<Drawing.BodyProperties>();
                    if (bodyPr != null)
                    {
                        // CONSISTENCY(spacing-units): accept unit-qualified
                        // input ('14pt', '0.5cm', '0.2in') and Get's 4-CSV
                        // 'Lpt,Tpt,Rpt,Bpt' readback for round-trip.
                        var (lE, tE, rE, bE) = ParseShapeMarginToEmu(value);
                        bodyPr.LeftInset = lE;
                        bodyPr.TopInset = tE;
                        bodyPr.RightInset = rE;
                        bodyPr.BottomInset = bE;
                    }
                    break;
                }
                case "preset" or "geometry" or "shape":
                {
                    // CONSISTENCY(shape-preset): mirror Add — replace prstGeom on
                    // ShapeProperties with the new preset token.
                    var spPr = shape.ShapeProperties;
                    if (spPr != null)
                    {
                        var newPreset = ParseExcelShapePreset(value);
                        spPr.RemoveAllChildren<Drawing.PresetGeometry>();
                        spPr.AppendChild(new Drawing.PresetGeometry(new Drawing.AdjustValueList()) { Preset = newPreset });
                    }
                    break;
                }
                default:
                    shpUnsupported.Add(key);
                    break;
            }
        }

        drawingsPart.WorksheetDrawing.Save();
        return shpUnsupported;
    }

    private List<string> SetSlicerByPath(Match m, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var slIdx = int.Parse(m.Groups[1].Value);
        if (!TryFindSlicerByIndex(worksheet, slIdx, out var slicer, out _) || slicer == null)
            throw new ArgumentException($"slicer[{slIdx}] not found on sheet");

        var slicersPart = worksheet.GetPartsOfType<SlicersPart>().FirstOrDefault();
        var slUnsupported = new List<string>();
        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "caption": slicer.Caption = value; break;
                case "style": slicer.Style = value; break;
                case "name": slicer.Name = value; break;
                case "rowheight":
                    if (uint.TryParse(value, out var rh)) slicer.RowHeight = rh;
                    else slUnsupported.Add(key);
                    break;
                case "columncount":
                    if (uint.TryParse(value, out var cc) && cc >= 1 && cc <= 20000)
                        slicer.ColumnCount = cc;
                    else slUnsupported.Add(key);
                    break;
                default: slUnsupported.Add(key); break;
            }
        }
        if (slicersPart?.Slicers != null) slicersPart.Slicers.Save(slicersPart);
        SaveWorksheet(worksheet);
        return slUnsupported;
    }

    // CONSISTENCY(table-column-path): mirror the col[M].prop= dotted form already
    // accepted on /Sheet/table[N] by exposing the column as a sub-path so users
}
